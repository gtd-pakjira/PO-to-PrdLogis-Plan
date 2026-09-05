"""
grouping.py — Module 2: Grouping + Reconciliation

รวมยอดจาก po_line (ราย fc_code) เข้าเป็นยอดต่อ SKU x กลุ่มพื้นที่ (บางบัวทอง/มหาชัย/สุวรรณภูมิ)
ผ่านตาราง location_mapping แล้วตรวจสอบว่ายอดรวมตรงกับ PO ต้นฉบับ 100% (UC-9 Reconciliation)

วิธีรัน (จาก root ของโปรเจกต์):
    python -m src.grouping <po_import_id>
"""
import sys

from customers.cpall.logic.db import get_connection


class ReconciliationError(Exception):
    pass


class InactiveSkuOrderedError(Exception):
    """PO รอบนี้สั่งสินค้าที่ถูกปิดใช้งาน (is_active=False) ใน Django Admin อยู่ — ต้องไป active
    สินค้านั้นก่อน ถึงจะสร้างแผนได้ (ยืนยันกับ user แล้วว่าต้อง block ไม่ใช่แค่เตือน)"""
    pass


def check_inactive_skus_ordered(po_import_ids) -> list[dict]:
    """
    หาบาร์โค้ดที่ PO รอบนี้สั่งอยู่จริง แต่ถูกปิดใช้งาน (is_active=False) ใน product_master —
    คืน [{"barcode":, "name_th":}, ...] ว่างเปล่าถ้าไม่มีปัญหา
    """
    from django.db.models import Q

    from customers.cpall.models import PoLine, ProductMaster

    if isinstance(po_import_ids, int):
        po_import_ids = [po_import_ids]

    ordered_barcodes = set(
        PoLine.objects.filter(po_import_id__in=po_import_ids).values_list("barcode", flat=True).distinct()
    )
    inactive = ProductMaster.objects.filter(
        Q(barcode__in=ordered_barcodes) & Q(is_active=False)
    ).values("barcode", "name_th")
    return list(inactive)


def get_grouped_quantities(po_import_ids) -> list[dict]:
    """
    คืนยอดสั่งรวมต่อ (barcode, group) สำหรับ po_import_id(s) ที่ระบุ (int เดี่ยว หรือ list)
    -> [{"barcode": ..., "group": ..., "qty_case_ordered": ...}, ...]
    """
    if isinstance(po_import_ids, int):
        po_import_ids = [po_import_ids]

    conn = get_connection()
    try:
        with conn.cursor() as cur:
            cur.execute(
                """
                SELECT pl.barcode, lm."group", SUM(pl.qty_ordered) AS qty
                FROM po_line pl
                JOIN location_mapping lm ON pl.fc_code = lm.fc_code
                WHERE pl.po_import_id = ANY(%s)
                GROUP BY pl.barcode, lm."group"
                ORDER BY pl.barcode, lm."group"
                """,
                (po_import_ids,),
            )
            rows = cur.fetchall()
    finally:
        conn.close()

    return [{"barcode": r[0], "group": r[1], "qty_case_ordered": float(r[2])} for r in rows]


def get_grouped_quantities_by_sub_location(po_import_ids) -> list[dict]:
    """
    เหมือน get_grouped_quantities แต่ละเอียดถึงระดับ sub_location (จุดส่งย่อย)
    ใช้สำหรับ Module 4 (Excel Exporter) ที่ต้องกรอกยอดแยกคอลัมน์ตามจุดส่งย่อยในเทมเพลต

    po_import_ids: รับได้ทั้ง int เดี่ยว หรือ list ของ int (รวมหลายรอบ PO เข้าด้วยกัน — เช่น
    Production Plan ที่ต้องรวมรอบบ่าย (บางบัวทอง/มหาชัย/สุวรรณภูมิ) กับรอบเช้า (รอบเช้าต่างจังหวัด))
    -> [{"barcode": ..., "sub_location": ..., "qty_case_ordered": ...}, ...]
    """
    if isinstance(po_import_ids, int):
        po_import_ids = [po_import_ids]

    conn = get_connection()
    try:
        with conn.cursor() as cur:
            cur.execute(
                """
                SELECT pl.barcode, lm.sub_location, SUM(pl.qty_ordered) AS qty
                FROM po_line pl
                JOIN location_mapping lm ON pl.fc_code = lm.fc_code
                WHERE pl.po_import_id = ANY(%s)
                GROUP BY pl.barcode, lm.sub_location
                ORDER BY pl.barcode, lm.sub_location
                """,
                (po_import_ids,),
            )
            rows = cur.fetchall()
    finally:
        conn.close()

    return [{"barcode": r[0], "sub_location": r[1], "qty_case_ordered": float(r[2])} for r in rows]


def get_grouped_quantities_by_sub_location_and_po(po_import_ids) -> list[dict]:
    """
    ละเอียดถึงระดับ (barcode, sub_location, po_number) — ใช้สำหรับ Logistic Plan
    ที่แยกยอดเป็นคอลัมน์ต่อ PO ภายในแต่ละจุดส่งย่อย (จุดส่งหนึ่งอาจมีหลาย PO ในรอบเดียวกัน)
    รับได้ทั้ง int เดี่ยว หรือ list ของ int
    -> [{"barcode":, "sub_location":, "po_number":, "qty_case_ordered":}, ...]
    """
    if isinstance(po_import_ids, int):
        po_import_ids = [po_import_ids]

    conn = get_connection()
    try:
        with conn.cursor() as cur:
            cur.execute(
                """
                SELECT pl.barcode, lm.sub_location, pl.po_number, SUM(pl.qty_ordered) AS qty
                FROM po_line pl
                JOIN location_mapping lm ON pl.fc_code = lm.fc_code
                WHERE pl.po_import_id = ANY(%s)
                GROUP BY pl.barcode, lm.sub_location, pl.po_number
                ORDER BY lm.sub_location, pl.po_number, pl.barcode
                """,
                (po_import_ids,),
            )
            rows = cur.fetchall()
    finally:
        conn.close()

    return [
        {"barcode": r[0], "sub_location": r[1], "po_number": r[2], "qty_case_ordered": float(r[3])}
        for r in rows
    ]


def get_covered_sub_locations(po_import_ids) -> set:
    """
    คืนชุด sub_location ที่ "มีข้อมูลจริง" ใน po_import_id(s) ที่ระบุ (ไม่ว่า SKU ไหน)
    ใช้แยกแยะ "จุดนี้มีรอบ PO มาแล้วแต่ SKU นี้สั่ง 0" (ต้องกรอก 0) ออกจาก
    "จุดนี้ยังไม่มีรอบ PO มาเลย" (ต้องเว้นว่างไว้ ไม่ใช่ 0)
    """
    if isinstance(po_import_ids, int):
        po_import_ids = [po_import_ids]

    conn = get_connection()
    try:
        with conn.cursor() as cur:
            cur.execute(
                """
                SELECT DISTINCT lm.sub_location
                FROM po_line pl
                JOIN location_mapping lm ON pl.fc_code = lm.fc_code
                WHERE pl.po_import_id = ANY(%s)
                """,
                (po_import_ids,),
            )
            return {r[0] for r in cur.fetchall()}
    finally:
        conn.close()


def get_dates_for_po_import(po_import_id: int):
    """คืนค่า (production_date, po_date) ที่ผูกไว้กับ po_import_id นี้ตอน import"""
    conn = get_connection()
    try:
        with conn.cursor() as cur:
            cur.execute("SELECT production_date, po_date FROM po_import WHERE id = %s", (po_import_id,))
            row = cur.fetchone()
            return (row[0], row[1]) if row else (None, None)
    finally:
        conn.close()


def get_dates_by_sub_location(po_import_ids) -> dict:
    """
    คืน {sub_location: (production_date, po_date)} — หาว่าจุดส่งย่อยแต่ละจุดมีข้อมูลอยู่ใน po_import_id
    ตัวไหน (ในบรรดาที่ระบุมา) แล้วดึงวันที่ของรอบนั้นมาผูกให้ ใช้ตอนสร้างไฟล์ที่รวมหลายรอบเข้าด้วยกัน
    (เช่น Production Plan) ที่แต่ละคอลัมน์อาจต้องโชว์วันที่คนละชุดกัน เพราะมาจากคนละรอบ PO
    """
    if isinstance(po_import_ids, int):
        po_import_ids = [po_import_ids]

    conn = get_connection()
    try:
        with conn.cursor() as cur:
            cur.execute(
                """
                SELECT DISTINCT lm.sub_location, pl.po_import_id
                FROM po_line pl
                JOIN location_mapping lm ON pl.fc_code = lm.fc_code
                WHERE pl.po_import_id = ANY(%s)
                """,
                (po_import_ids,),
            )
            sub_location_to_import_id = {r[0]: r[1] for r in cur.fetchall()}
    finally:
        conn.close()

    result = {}
    for sub_loc, import_id in sub_location_to_import_id.items():
        result[sub_loc] = get_dates_for_po_import(import_id)
    return result


def reconcile(po_import_ids) -> dict:
    """
    เทียบยอดรวมที่ group แล้ว กับยอดรวมดิบใน po_line ต่อ barcode
    ต้องตรงกัน 100% ไม่งั้น raise ReconciliationError (ตาม NFR-1)
    รับได้ทั้ง int เดี่ยว หรือ list ของ int (เผื่อเช็ครวมหลายรอบ PO พร้อมกัน)

    หมายเหตุ: ถ้ามี fc_code ที่ไม่อยู่ใน location_mapping แถวนั้นจะหายไปจาก JOIN
    (เป็นสาเหตุที่พบบ่อยที่สุดของยอดไม่ตรง) — ให้รัน check_unknown_locations()
    ใน po_parser.py ก่อนเสมอ
    """
    if isinstance(po_import_ids, int):
        po_import_ids = [po_import_ids]

    grouped = get_grouped_quantities(po_import_ids)
    grouped_total_by_sku = {}
    for row in grouped:
        grouped_total_by_sku.setdefault(row["barcode"], 0.0)
        grouped_total_by_sku[row["barcode"]] += row["qty_case_ordered"]

    conn = get_connection()
    try:
        with conn.cursor() as cur:
            cur.execute(
                """
                SELECT barcode, SUM(qty_ordered)
                FROM po_line
                WHERE po_import_id = ANY(%s)
                GROUP BY barcode
                """,
                (po_import_ids,),
            )
            raw_total_by_sku = {r[0]: float(r[1]) for r in cur.fetchall()}
    finally:
        conn.close()

    mismatches = []
    for barcode, raw_qty in raw_total_by_sku.items():
        grouped_qty = grouped_total_by_sku.get(barcode, 0.0)
        if abs(raw_qty - grouped_qty) > 0.001:
            mismatches.append(
                {"barcode": barcode, "po_total": raw_qty, "grouped_total": grouped_qty}
            )

    result = {
        "po_import_ids": po_import_ids,
        "sku_count": len(raw_total_by_sku),
        "po_grand_total": sum(raw_total_by_sku.values()),
        "grouped_grand_total": sum(grouped_total_by_sku.values()),
        "mismatches": mismatches,
        "passed": len(mismatches) == 0,
    }
    return result


if __name__ == "__main__":
    if len(sys.argv) < 2:
        print("Usage: python -m src.grouping <po_import_id> [<po_import_id_2> ...]")
        sys.exit(1)

    po_import_ids = [int(x) for x in sys.argv[1:]]
    result = reconcile(po_import_ids)

    print(f"[grouping] SKU ทั้งหมด: {result['sku_count']}")
    print(f"[grouping] ยอดรวมจาก PO ดิบ:     {result['po_grand_total']:,.2f} ลัง")
    print(f"[grouping] ยอดรวมหลัง group:     {result['grouped_grand_total']:,.2f} ลัง")

    if result["passed"]:
        print("[grouping] ✅ Reconciliation ผ่าน — ยอดตรงกัน 100%")
    else:
        print(f"[grouping] ❌ พบยอดไม่ตรงกัน {len(result['mismatches'])} SKU:")
        for m in result["mismatches"]:
            print(f"    - {m['barcode']}: PO={m['po_total']} vs grouped={m['grouped_total']}")
        raise ReconciliationError("ยอดไม่ตรงกัน ต้องแก้ไขก่อนไปทำ Production Plan")
