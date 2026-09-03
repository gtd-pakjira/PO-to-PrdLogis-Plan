"""
po_regenerator.py — สร้างไฟล์ .xlsx ใหม่จากข้อมูลที่เก็บไว้ใน database ครบทุกคอลัมน์เหมือนไฟล์ต้นฉบับ
(ไม่ใช่แค่ 12 คอลัมน์ที่ระบบใช้คำนวณจริง) — ใช้เป็น fallback ตอนกด "ดาวน์โหลด" ไฟล์ PO ที่เคย import
ไว้แล้ว แต่ไฟล์ต้นฉบับหายไปจากที่เก็บ (ตามปกติแล้วจะเสิร์ฟไฟล์ต้นฉบับจริงตรงๆ ก่อนเสมอ — ดู
views.py's download_po — ฟังก์ชันนี้เป็นแค่ทางสำรอง)

เก็บ+เขียนกลับด้วย "ตำแหน่ง" ไม่ใช่ "ชื่อคอลัมน์" เพราะไฟล์จริงของ CP All มีชื่อคอลัมน์ซ้ำกันได้ (เช่น
"Discount Percentage 1" ปรากฏ 2 รอบในไฟล์เดียว) — PoImport.column_order (ลำดับชื่อคอลัมน์ต้นฉบับ) กับ
PoLine.all_values (ค่าตามลำดับเดียวกัน) ถูกเก็บไว้ตอน import แล้ว (ดู po_parser.py's parse_po_file)
"""
import io

import openpyxl


class PORegenerateError(Exception):
    pass


def regenerate_po_file_bytes(po_import_id: int) -> bytes:
    from customers.cpall.models import PoImport, PoLine

    try:
        po_import = PoImport.objects.get(id=po_import_id)
    except PoImport.DoesNotExist:
        raise PORegenerateError("ไม่พบ PO นี้")

    if not po_import.column_order:
        raise PORegenerateError(
            "PO นี้ไม่มีข้อมูล column_order เก็บไว้ (อาจเป็น PO เก่าก่อนมีระบบนี้) — สร้างไฟล์ใหม่แบบครบทุกคอลัมน์ไม่ได้"
        )

    lines = PoLine.objects.filter(po_import_id=po_import_id).order_by("id")

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "po_export"
    ws.append(po_import.column_order)

    for ln in lines:
        ws.append(ln.all_values if ln.all_values else [])

    buffer = io.BytesIO()
    wb.save(buffer)
    return buffer.getvalue()
