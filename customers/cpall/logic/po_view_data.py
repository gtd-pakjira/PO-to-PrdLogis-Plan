"""
po_view_data.py — อ่านข้อมูล PO ที่เคย import ไว้ กลับมาแสดงเป็นตารางในหน้าเว็บ (ไม่ผ่านไฟล์เลย
อ่านจาก po_line ตรงๆ เร็วกว่า ค้นหาง่ายกว่า) — po_line เก็บทุกคอลัมน์ที่จำเป็นไว้ครบตั้งแต่ตอน import
(ตรงกับ REQUIRED_COLUMNS ของ po_parser.py แบบ 1:1)

หมายเหตุ: ไฟล์ .xlsx ต้นฉบับยังคงเก็บไว้ถาวรบนดิสก์ (ตัดสินใจร่วมกับผู้ใช้แล้วว่าต้องดูย้อนหลังได้ว่า
นำเข้าอะไรเข้าระบบไปจริงๆ แบบเป๊ะๆ — เป็น audit trail) หน้านี้ (ตาราง) ใช้ข้อมูลจาก DB เพื่อความเร็ว/
ค้นหาง่ายเท่านั้น ส่วนดาวน์โหลดจริง (views.py's download_po) เสิร์ฟไฟล์ต้นฉบับตรงๆ ไม่ใช่ข้อมูลจากตรงนี้
"""
import os

from customers.cpall.models import PoImport, PoLine


def get_po_detail(po_import_id: int) -> dict | None:
    """ข้อมูลสรุป + รายการ SKU ทั้งหมดของ PO ที่ import ไว้ 1 รายการ — ใช้แสดงหน้าเว็บ"""
    try:
        po_import = PoImport.objects.get(id=po_import_id)
    except PoImport.DoesNotExist:
        return None

    lines = PoLine.objects.filter(po_import_id=po_import_id).order_by("id")

    return {
        "id": po_import.id,
        "source_filename": po_import.source_filename,  # path จริงบนดิสก์ — ใช้เปิด/ดาวน์โหลดไฟล์จริง
        "display_filename": os.path.basename(po_import.source_filename),  # ตัด path ไว้แสดงหน้าเว็บ
        "production_date": po_import.production_date,
        "po_date": po_import.po_date,
        "total_rows": po_import.total_rows,
        "imported_at": po_import.imported_at,
        "imported_by": po_import.imported_by,
        "lines": [
            {
                "po_number": ln.po_number, "po_date": ln.po_date, "delivery_date": ln.delivery_date,
                "delivery_time": ln.delivery_time, "fc_code": ln.fc_code,
                "delivery_location": ln.delivery_location, "line_no": ln.line_no,
                "barcode": ln.barcode, "item_name": ln.item_name, "qty_ordered": ln.qty_ordered,
                "unit_type": ln.unit_type, "net_case_price": ln.net_case_price,
                "total_amount": ln.total_amount,
            }
            for ln in lines
        ],
    }
