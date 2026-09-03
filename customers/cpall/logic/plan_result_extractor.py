"""
plan_result_extractor.py — Phase 1.6 sub-phase 3: ดึงผลลัพธ์ต่อ SKU/คอลัมน์จากไฟล์ Excel ที่สร้างแล้ว
มาเก็บลง plan_sku_result โดยใช้ค่าที่ LibreOffice คำนวณจริง (ไม่ใช่สูตรจำลองใน Python อีกต่อไป)

*** ทำไมต้องมีไฟล์นี้แยกจาก plan_view_data.py ***
plan_view_data.py (เดิม) อ่านไฟล์ openpyxl-only (ไม่เคยผ่านโปรแกรม spreadsheet จริง) จึงต้อง "จำลอง"
สูตรพวกนี้เป็น Python เอง (_pack_breakdown_text, _basket_total) — ไฟล์นี้อ่านจากไฟล์ที่ผ่าน
LibreOffice คำนวณมาแล้ว (excel_calc.load_calculated_workbook) จึงอ่านค่าที่คำนวณจริงได้ตรงๆ เลย
ไม่ต้องเขียนสูตรซ้ำเป็น Python — plan_view_data.py เดิมยังคงอยู่ (ใช้กับไฟล์ที่ยังไม่ผ่านขั้นตอนนี้
หรือ fallback) แต่ path ใหม่ (การสร้างแผน) จะใช้ไฟล์นี้แทน

*** ข้อจำกัดที่ยอมรับไว้ก่อนในรอบนี้ ***
basket_total (รวมตะกร้า) อยู่คนละชีตกับข้อมูลหลัก ("-รถ") และอ้างอิง SKU ด้วยชื่อ+ตำแหน่งแถว ไม่ใช่
บาร์โค้ดตรงๆ — การจับคู่ข้ามชีตแบบเชื่อถือได้ 100% ยังไม่ได้ทำในรอบนี้ (เสี่ยงจับคู่ผิดแถวถ้าเทมเพลต
มีโครงสร้างไม่ตรงกันเป๊ะ) เลย "ยืม" ใช้ _basket_total() ที่ทดสอบแล้วว่าคำนวณตรงกับ LibreOffice เป๊ะ
(ดู test_plan_view_data.py) ไปพลางก่อน — จุดนี้บันทึกไว้ชัดเจนว่าเป็นการลดขอบเขตที่ตั้งใจ ไม่ใช่มองข้าม
"""
from customers.cpall.logic.excel_calc import load_calculated_workbook
from customers.cpall.logic.excel_export import BUFFER_COL, BUFFER_ROW_OFFSET, _find_sub_location_columns
from customers.cpall.logic.excel_export import COL_NAME as PP_COL_NAME
from customers.cpall.logic.excel_export import SHEET_NAME as PP_SHEET_NAME
from customers.cpall.logic.excel_export import _find_sku_header_rows as _find_pp_sku_header_rows
from customers.cpall.logic.logistic_plan_export import (
    _find_column_labels,
    _find_line_no_column,
    _find_qty_column_range,
    get_group_templates,
)
from customers.cpall.logic.logistic_plan_export import _find_sku_header_rows as _find_lp_sku_header_rows
from customers.cpall.logic.plan_view_data import _basket_total


def _format_pack_text(value):
    """
    แปลงค่า 'แพค/เศษ' ที่ LibreOffice คำนวณจากสูตรของ Admin ให้เป็นข้อความเสมอ

    สูตรจริงในเทมเพลต (เช่น "IF(qty=0,"", IF(MOD(qty,pack)=0, qty/pack, ...))") ตั้งใจให้คืนค่า
    เป็น "ตัวเลขล้วน" (ไม่ใช่ string) เมื่อยอดสั่งหารด้วยขนาดแพคลงตัวพอดี (ไม่มีเศษ) — เดิมโค้ดนี้เช็ค
    isinstance(value, str) แล้วทิ้งเป็น None ทันทีถ้าไม่ใช่ string ทำให้กรณี "ลงตัวพอดี" (ซึ่งเกิดขึ้น
    บ่อยมาก ไม่ใช่ edge case แปลกๆ) แสดงเป็น "None" ในหน้าเว็บทั้งที่ยอดสั่งจริงถูกต้องอยู่แล้ว — เป็น
    บั๊กจริงที่พบจากการทดสอบจริง ไม่ใช่แค่การแสดงผล (pack_text ที่เก็บใน DB เองก็เป็น None ผิดจริง)
    """
    if value is None:
        return None
    if isinstance(value, str):
        return value if value != "" else None  # สูตรคืน "" ตอนยอดสั่ง=0 -> ไม่มีอะไรให้แสดง
    if isinstance(value, (int, float)):
        return str(int(value)) if value == int(value) else str(value)
    return str(value)


PP_COL_PRICE = 5
PP_COL_PACK = 6
PP_NAME_EN_ROW_OFFSET = 1
PP_RETURN_ROW_OFFSET = 3


def extract_production_plan_results(filepath: str) -> list[dict]:
    """
    เปิดไฟล์ Production Plan ผ่าน LibreOffice คำนวณสูตรจริงก่อน แล้วดึงผลลัพธ์ต่อ SKU/จุดส่งย่อยออกมา
    เป็น list ของ dict พร้อมเก็บลง plan_sku_result (sheet_type='production')
    """
    wb = load_calculated_workbook(filepath)
    ws = wb[PP_SHEET_NAME]

    col_to_sub_location = _find_sub_location_columns(ws)
    header_rows = _find_pp_sku_header_rows(ws)

    rows = []
    for barcode, row in sorted(header_rows.items(), key=lambda kv: kv[1]):
        name_th = ws.cell(row=row, column=PP_COL_NAME).value
        name_en = ws.cell(row=row + PP_NAME_EN_ROW_OFFSET, column=PP_COL_NAME).value
        price = ws.cell(row=row, column=PP_COL_PRICE).value
        pack_size = ws.cell(row=row, column=PP_COL_PACK).value
        buffer_qty = ws.cell(row=row + BUFFER_ROW_OFFSET, column=BUFFER_COL).value
        return_qty = ws.cell(row=row + PP_RETURN_ROW_OFFSET, column=BUFFER_COL).value  # ค่าจริงจาก LibreOffice

        qty_by_location = {}
        grand_total = 0
        for col, sub_loc in col_to_sub_location.items():
            val = ws.cell(row=row, column=col).value
            qty_by_location[sub_loc] = val
            if isinstance(val, (int, float)):
                grand_total += val
            pack_text = ws.cell(row=row + PP_NAME_EN_ROW_OFFSET, column=col).value  # ค่าจริงจาก LibreOffice
            rows.append({
                "sheet_type": "production", "group_name": None, "barcode": barcode,
                "name_th": name_th, "name_en": name_en, "price": price, "pack_size": pack_size,
                "column_label": sub_loc, "qty": val,
                "pack_text": _format_pack_text(pack_text),
                "grand_total": None,  # เติมทีหลังหลังรวมยอดครบทุกคอลัมน์ของ SKU นี้
                "buffer_qty": buffer_qty, "return_qty": return_qty, "basket_total": None,
            })

        for r in rows:
            if r["barcode"] == barcode and r["sheet_type"] == "production":
                r["grand_total"] = grand_total

    return rows


def extract_logistic_plan_results(filepath: str, group_name: str) -> list[dict]:
    """เหมือน extract_production_plan_results แต่สำหรับ Logistic Plan 1 กลุ่ม"""
    _, sheet_name = get_group_templates()[group_name]
    wb = load_calculated_workbook(filepath)
    ws = wb[sheet_name]

    line_no_col, header_row = _find_line_no_column(ws)
    name_col = line_no_col + 1
    pack_col = line_no_col + 2
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

    col_to_label = {}
    for col, (sub_loc, po_idx) in col_labels.items():
        col_to_label[col] = f"{sub_loc} PO{po_idx}" if po_idx else sub_loc

    header_rows = _find_lp_sku_header_rows(ws, name_col)

    rows = []
    for barcode, row in sorted(header_rows.items(), key=lambda kv: kv[1]):
        name_th = ws.cell(row=row, column=name_col).value
        raw_name_en = ws.cell(row=row + 1, column=name_col).value
        name_en = raw_name_en
        if isinstance(raw_name_en, str):
            import re
            name_en = re.sub(r"\(\d{10,16}\)\s*$", "", raw_name_en).strip()
        pack_size = ws.cell(row=row, column=pack_col).value

        qty_by_column = {}
        grand_total = 0
        for col, label in col_to_label.items():
            val = ws.cell(row=row, column=col).value
            qty_by_column[label] = val
            if isinstance(val, (int, float)):
                grand_total += val

        # basket_total: ยังใช้ฟังก์ชัน Python ที่ทดสอบแล้วว่าตรงกับ LibreOffice เป๊ะไปก่อน (ดู docstring
        # บนสุดของไฟล์ — เหตุผลที่ยังไม่อ่านจากชีต "-รถ" ตรงๆ ในรอบนี้)
        basket_total = _basket_total(qty_by_column, pack_size)

        for col, label in col_to_label.items():
            val = ws.cell(row=row, column=col).value
            pack_text = ws.cell(row=row + 1, column=col).value  # ค่าจริงจาก LibreOffice
            rows.append({
                "sheet_type": "logistic", "group_name": group_name, "barcode": barcode,
                "name_th": name_th, "name_en": name_en, "price": None, "pack_size": pack_size,
                "column_label": label, "qty": val,
                "pack_text": _format_pack_text(pack_text),
                "grand_total": grand_total, "buffer_qty": None, "return_qty": None,
                "basket_total": basket_total,
            })

    return rows
