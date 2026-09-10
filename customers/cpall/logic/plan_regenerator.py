"""
plan_regenerator.py — Phase 1.6 sub-phase 4: สร้างไฟล์ Excel ใหม่ตอนกด "ดาวน์โหลด" จากข้อมูลดิบ
(plan_sku_result) + เทมเพลตเวอร์ชันที่ผูกไว้กับแผนนั้น (production_template_version / template_version)

*** หลักการสำคัญ: ไม่คำนวณอะไรเองเลย เขียนแค่ "ค่าดิบ" ที่ Admin กรอก/ระบบคำนวณจาก PO เท่านั้น ***
(ยอดสั่งต่อ SKU/คอลัมน์, ยอดเผื่อ) ส่วนสูตรทั้งหมด (ยอดรวม, ยอดคืน, แตกลัง/เศษ) "ปล่อยทิ้งไว้เป็นสูตร"
ในเทมเพลตต่อไป ไม่แตะเลย — ไฟล์ที่ได้จะมีสูตรจริงครบเหมือนที่ Admin คุ้นเคย เปิดด้วย Excel/LibreOffice
จริงแล้วคำนวณเองอัตโนมัติ (ตรงกับที่ยืนยันไว้ว่าต้องมีสูตรอยู่ในไฟล์ที่ดาวน์โหลด)

ทำงานได้เฉพาะแผนที่สร้างหลัง Phase 1.6 sub-phase 3+5 (มีข้อมูลใน plan_sku_result และผูก template
version ไว้แล้ว) — แผนเก่ากว่านั้น raise PlanRegenerateError ให้ผู้เรียก fallback ไปเสิร์ฟไฟล์เดิมที่
ยังเก็บไว้บนดิสก์แทน (ไฟล์เก่ายังไม่ถูกลบในเฟสนี้ — เก็บไว้เป็นทางสำรองก่อน)
"""
import io

import openpyxl

# from customers.cpall.logic.date_utils import find_merged_date_header_column, update_date_headers
from customers.cpall.logic.date_utils import update_date_headers
# from customers.cpall.logic.excel_export import BUFFER_COL, BUFFER_ROW_OFFSET, _find_sub_location_columns
# from customers.cpall.logic.excel_export import _find_sku_header_rows as _find_pp_sku_header_rows
from customers.cpall.logic.excel_export import (
    BUFFER_COL,
    BUFFER_ROW_OFFSET,
    _find_sub_location_columns,
    _find_sku_header_rows as _find_pp_sku_header_rows,
    _renumber_visible_sku_rows as _renumber_pp_visible_sku_rows,
)
from customers.cpall.logic.excel_export import get_sheet_name as get_pp_sheet_name
# from customers.cpall.logic.grouping import get_dates_by_sub_location
from customers.cpall.logic.grouping import (
    get_plan_date_context,
    MORNING_GROUP_NAME,
)
from customers.cpall.logic.logistic_plan_export import (
    _find_column_labels,
    _find_line_no_column,
    _find_qty_column_range,
    _renumber_logistic_sku_rows,
    get_group_templates,
    _find_buffer_column,
)
from customers.cpall.logic.logistic_plan_export import _find_sku_header_rows as _find_lp_sku_header_rows


class PlanRegenerateError(Exception):
    pass


def _update_dates(ws, plan_run, col_to_sub_location):
    """
    อัปเดตวันที่ใน Production Plan จาก Date Context เดียวกับตอนสร้างแผน

    แต่ละ sub_location จะเลือก context ตามกลุ่ม:
    - รอบเช้าต่างจังหวัด → morning
    - กลุ่มอื่น → afternoon

    Date Context รองรับทั้ง actual date และ fallback date
    ตาม business rule ใน grouping.get_plan_date_context()
    """
    po_import_ids = list(plan_run.po_imports.values_list("id", flat=True))

    if not po_import_ids:
        return

    date_context = get_plan_date_context(po_import_ids)

    from customers.cpall.models import LocationMapping

    sub_locations = set(col_to_sub_location.values())

    group_by_sub_location = dict(
        LocationMapping.objects
        .filter(sub_location__in=sub_locations)
        .values_list("sub_location", "group")
    )

    def date_resolver(col):
        sub_loc = col_to_sub_location.get(col)

        if sub_loc is None:
            return None

        group_name = group_by_sub_location.get(sub_loc)

        if group_name == MORNING_GROUP_NAME:
            return date_context["morning"]

        return date_context["afternoon"]

    n = update_date_headers(ws, date_resolver)

    print(
        f"[plan_regenerator] อัปเดตวันที่ในหัวไฟล์ {n} จุด "
        f"(ใช้ Date Context เดียวกับตอนสร้างแผน)"
    )


def regenerate_production_plan_bytes(plan_run_id: int) -> bytes:
    """สร้างไฟล์ Production Plan ใหม่จากข้อมูลดิบ คืนเป็น bytes ตรงๆ (ไม่ผ่านไฟล์ชั่วคราวบนดิสก์เลย)"""
    from customers.cpall.models import PlanRun, PlanSkuResult

    try:
        plan_run = PlanRun.objects.get(id=plan_run_id)
    except PlanRun.DoesNotExist:
        raise PlanRegenerateError("ไม่พบแผนนี้")

    if plan_run.production_template_version is None:
        raise PlanRegenerateError("แผนนี้ไม่ได้ผูกเทมเพลตเวอร์ชันไว้ (อาจเป็นแผนเก่าก่อนมีระบบนี้)")

    results = list(PlanSkuResult.objects.filter(plan_run_id=plan_run_id, sheet_type="production"))
    if not results:
        raise PlanRegenerateError("ไม่มีข้อมูลดิบของแผนนี้ (อาจเป็นแผนเก่าก่อนมีระบบนี้)")

    by_barcode = {}
    for r in results:
        entry = by_barcode.setdefault(
            r.barcode, {"buffer_qty": r.buffer_qty, "qty_by_col": {}, "grand_total": r.grand_total}
        )
        entry["qty_by_col"][r.column_label] = r.qty

    wb = openpyxl.load_workbook(plan_run.production_template_version.file_path)
    ws = wb[get_pp_sheet_name()]

    col_to_sub_location = _find_sub_location_columns(ws)
    sub_location_to_col = {v: k for k, v in col_to_sub_location.items()}
    header_rows = _find_pp_sku_header_rows(ws)

    # ... เขียน quantity / buffer ...
    for barcode, row in header_rows.items():
        data = by_barcode.get(barcode)
        if data is None:
            continue  # SKU ในเทมเพลตแต่ไม่มีข้อมูล (ไม่เคยสั่งรอบนี้) -> ปล่อยว่างไว้เหมือนตอนสร้างแผนจริง
        for sub_loc, qty in data["qty_by_col"].items():
            col = sub_location_to_col.get(sub_loc)
            if col is not None and qty is not None:
                ws.cell(row=row, column=col, value=float(qty))
        if data["buffer_qty"] is not None:
            ws.cell(row=row + BUFFER_ROW_OFFSET, column=BUFFER_COL, value=float(data["buffer_qty"]))

    # ... ซ่อน inactive SKU ...
    # สินค้าที่ปิดใช้งาน (is_active=False) และไม่มี PO สั่งเลยในรอบนี้ แต่ยังมีแถวอยู่ในเทมเพลต — ซ่อน
    # แถวไว้ (ไม่ลบจริง) เหมือนกับตอนสร้างแผนครั้งแรก (ดู excel_export.py) — *** เจอบั๊กจริง
    # (2025-09-05): ตอนสร้างแผนครั้งแรกซ่อนถูกต้อง แต่ตอนดาวน์โหลดซ้ำ/regenerate (ฟังก์ชันนี้) ไม่เคยมี
    # logic นี้เลย ทำให้ไฟล์ที่ Admin ดาวน์โหลดจริง (ผ่านฟังก์ชันนี้เสมอ ไม่ใช่ไฟล์ตอนสร้างแผนที่ถูกลบ
    # ทิ้งไปแล้วตาม data-first) ยังโผล่แถวสินค้าปิดใช้งานเป็นแถวปกติพร้อมค่า 0 — เหมือนบั๊ก M5 เป๊ะ
    # (แก้จุดสร้างแผนแต่ลืมแก้จุด regenerate) ***
    from customers.cpall.models import ProductMaster
    inactive_barcodes = set(
        ProductMaster.objects.filter(is_active=False).values_list("barcode", flat=True)
    )
    hidden_count = 0
    for barcode, row in header_rows.items():
        data = by_barcode.get(barcode)
        # "ไม่มี PO สั่งเลย" ต้องเช็คจาก grand_total (ยอดรวมจริงจาก PO ต้นทาง) ไม่ใช่แค่ "ไม่มีแถวใน
        # by_barcode" เพราะ by_barcode มาจาก PlanSkuResult ที่มีแถวของทุก SKU ในเทมเพลตอยู่แล้วเสมอ
        # (extraction insert ให้ครบทุก SKU ไม่ว่าจะมีคำสั่งซื้อจริงหรือไม่) — เจอบั๊กนี้ตอนเขียนโค้ดนี้
        # เอง (เช็คผิดเป็น "not in by_barcode" ซึ่งเป็น False เสมอ ไม่เคยซ่อนอะไรเลย)
        no_order = data is None or not data["grand_total"]
        if barcode in inactive_barcodes and no_order:
            for offset in range(4):
                ws.row_dimensions[row + offset].hidden = True
            hidden_count += 1
    # if hidden_count:
    #     print(f"[plan_regenerator] ซ่อน {hidden_count} SKU ที่ปิดใช้งานและไม่มี PO สั่งในรอบนี้ (4 แถวต่อ SKU)")

    # dates_by_sub_location = _update_dates(ws, plan_run, col_to_sub_location)
    if hidden_count:
        print(
            f"[plan_regenerator] ซ่อน {hidden_count} SKU "
            f"ที่ปิดใช้งานและไม่มี PO สั่งในรอบนี้ (4 แถวต่อ SKU)"
        )

    # ต้องจัดเลขใหม่หลังจากซ่อน SKU แล้ว
    renumbered_count = _renumber_pp_visible_sku_rows(ws, header_rows)

    print(
        f"[plan_regenerator] จัดเลขลำดับ Production ใหม่แล้ว "
        f"{renumbered_count} SKU"
    )

    _update_dates(
        ws,
        plan_run,
        col_to_sub_location,
    )

    buffer = io.BytesIO()
    wb.save(buffer)
    return buffer.getvalue()


def regenerate_logistic_plan_bytes(plan_run_id: int, group_name: str) -> bytes:
    """สร้างไฟล์ Logistic Plan ใหม่จากข้อมูลดิบสำหรับ 1 กลุ่ม คืนเป็น bytes ตรงๆ"""
    from customers.cpall.models import PlanRun, PlanRunLogisticFile, PlanSkuResult

    try:
        plan_run = PlanRun.objects.get(id=plan_run_id)
    except PlanRun.DoesNotExist:
        raise PlanRegenerateError("ไม่พบแผนนี้")

    logistic_file = PlanRunLogisticFile.objects.filter(
        plan_run_id=plan_run_id, group_name=group_name
    ).first()
    if logistic_file is None or logistic_file.template_version is None:
        raise PlanRegenerateError("กลุ่มนี้ของแผนไม่ได้ผูกเทมเพลตเวอร์ชันไว้ (อาจเป็นแผนเก่าก่อนมีระบบนี้)")

    results = list(PlanSkuResult.objects.filter(
        plan_run_id=plan_run_id, sheet_type="logistic", group_name=group_name
    ))
    if not results:
        raise PlanRegenerateError("ไม่มีข้อมูลดิบของกลุ่มนี้ (อาจเป็นแผนเก่าก่อนมีระบบนี้)")

    by_barcode = {}
    for r in results:
        entry = by_barcode.setdefault(r.barcode, {})
        entry[r.column_label] = r.qty

        production_buffer_rows = PlanSkuResult.objects.filter(
        plan_run_id=plan_run_id,
        sheet_type="production",
        buffer_qty__isnull=False,
    ).values(
        "barcode",
        "buffer_qty",
    )

    buffer_by_barcode = {
        row["barcode"]: float(row["buffer_qty"])
        for row in production_buffer_rows
        if row["buffer_qty"] is not None
    }
    
    wb = openpyxl.load_workbook(logistic_file.template_version.file_path)
    _, sheet_name = get_group_templates()[group_name]
    ws = wb[sheet_name]

    line_no_col, header_row = _find_line_no_column(ws)
    name_col = line_no_col + 1
    qty_start_col = line_no_col + 3
    qty_start_col, qty_end_col = _find_qty_column_range(ws, qty_start_col, header_row)

    col_labels = {}
    last_sub_location = None
    for col in range(qty_start_col, qty_end_col + 1):
        sub_loc, po_idx = _find_column_labels(ws, col, header_row)
        if sub_loc is None:
            sub_loc = last_sub_location if last_sub_location is not None else group_name
        last_sub_location = sub_loc
        col_labels[col] = (sub_loc, po_idx)
    if len(col_labels) == 1:
        only_col = list(col_labels.keys())[0]
        sub_loc, po_idx = col_labels[only_col]
        if po_idx is None:
            col_labels[only_col] = (sub_loc, 1)

    label_to_col = {}
    for col, (sub_loc, po_idx) in col_labels.items():
        label = f"{sub_loc} PO{po_idx}" if po_idx else sub_loc
        label_to_col[label] = col

    header_rows = _find_lp_sku_header_rows(ws, name_col)

    # ---------- ยอดเผื่อ ----------
    # ตอนนี้มีเฉพาะ "รอบเช้าต่างจังหวัด"
    # ใช้ค่าจาก Production Plan result ซึ่งเป็น source of truth ของยอดเผื่อ
    if group_name == "รอบเช้าต่างจังหวัด":
        buffer_col = _find_buffer_column(ws)

        for barcode, row in header_rows.items():
            buffer_qty = buffer_by_barcode.get(barcode)

            # ใช้ .value = โดยตรงเพื่อให้ None ล้างค่าที่ค้างใน template ได้
            ws.cell(
                row=row,
                column=buffer_col,
            ).value = buffer_qty

    for barcode, row in header_rows.items():
        data = by_barcode.get(barcode)
        if data is None:
            continue
        for label, qty in data.items():
            col = label_to_col.get(label)
            if col is not None and qty is not None:
                ws.cell(row=row, column=col, value=float(qty))


    # สินค้าที่ปิดใช้งาน (is_active=False) และไม่มี PO สั่งเลยในกลุ่มนี้รอบนี้ แต่ยังมีแถวอยู่ในเทมเพลต
    # — ซ่อนแถวไว้เหมือนตอนสร้างแผนครั้งแรก (ดู logistic_plan_export.py) *** เจอบั๊กจริง (2025-09-05):
    # เหมือนกับ Production Plan เป๊ะ — ซ่อนถูกต้องตอนสร้างแผนครั้งแรก แต่ไม่เคยซ่อนเลยตอนดาวน์โหลดซ้ำ/
    # regenerate (ฟังก์ชันนี้) *** — Logistic Plan มีแค่ 2 แถวต่อ SKU (ต่างจาก Production Plan ที่มี 4)
    from customers.cpall.models import ProductMaster
    inactive_barcodes = set(
        ProductMaster.objects.filter(is_active=False).values_list("barcode", flat=True)
    )
    # hidden_count = 0
    # for barcode, row in header_rows.items():
    #     data = by_barcode.get(barcode)
    #     no_order = data is None or not any(v for v in data.values() if v)
    #     if barcode in inactive_barcodes and no_order:
    #         ws.row_dimensions[row].hidden = True
    #         ws.row_dimensions[row + 1].hidden = True
    #         hidden_count += 1
    # if hidden_count:
    #     print(f"[plan_regenerator:{group_name}] ซ่อน {hidden_count} SKU ที่ปิดใช้งานและไม่มี PO สั่งในรอบนี้")

    hidden_count = 0

    # ซ่อนเฉพาะ SKU ที่ inactive และไม่มี PO ในรอบนี้
    for barcode, row in header_rows.items():
        data = by_barcode.get(barcode)
        no_order = data is None or not any(v for v in data.values() if v)

        if barcode in inactive_barcodes and no_order:
            ws.row_dimensions[row].hidden = True
            ws.row_dimensions[row + 1].hidden = True
            hidden_count += 1

    if hidden_count:
        print(
            f"[plan_regenerator:{group_name}] "
            f"ซ่อน {hidden_count} SKU "
            f"ที่ปิดใช้งานและไม่มี PO สั่งในรอบนี้"
        )

    # SKU ที่มี PO หรือยัง active → จัดเลขลำดับใหม่
    ordered_barcodes = {
        barcode
        for barcode, data in by_barcode.items()
        if any(v for v in data.values() if v)
    }

    renumbered_count = _renumber_logistic_sku_rows(
        ws=ws,
        header_rows=header_rows,
        line_no_col=line_no_col,
        inactive_barcodes=inactive_barcodes,
        ordered_barcodes=ordered_barcodes,
    )

    print(
        f"[plan_regenerator:{group_name}] "
        f"จัดเลขลำดับใหม่แล้ว {renumbered_count} SKU"
    )

    col_to_sub_location = {col: sub_loc for col, (sub_loc, _) in col_labels.items()}
    _update_dates(ws, plan_run, col_to_sub_location)

    buffer = io.BytesIO()
    wb.save(buffer)
    return buffer.getvalue()
