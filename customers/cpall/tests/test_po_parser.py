"""
test_po_parser.py — เทส parse_po_file() (อ่าน+ตรวจสอบ+ตัดข้อมูลซ้ำ) ด้วยไฟล์ Excel จำลองที่สร้างขึ้น
เองในเทส (ไม่ใช้ไฟล์จริงของลูกค้า) ไม่แตะ database เลย (parse_po_file ไม่คุยกับ DB)
"""
import os
import tempfile

import openpyxl
from django.test import SimpleTestCase

from customers.cpall.logic.po_parser import REQUIRED_COLUMNS, POParseError, parse_po_file

# แถวตัวอย่าง 1 แถว เรียงตามลำดับ REQUIRED_COLUMNS เป๊ะ
SAMPLE_ROW = ["PO1", "27/08/2026", "28/08/2026", "10:00", "FC01", "คลังทดสอบ",
              1, "8859388000025", "องุ่นดำ", 10, "CT", 21.5]


def _make_po_excel(rows, omit_column=None):
    """สร้างไฟล์ .xlsx ชั่วคราวตามคอลัมน์ที่ระบบต้องการ คืน path ให้ (ลบเองหลังใช้)"""
    wb = openpyxl.Workbook()
    ws = wb.active
    headers = [c for c in REQUIRED_COLUMNS if c != omit_column]
    ws.append(headers)
    for row in rows:
        ws.append(row)
    fd, path = tempfile.mkstemp(suffix=".xlsx")
    os.close(fd)
    wb.save(path)
    return path


class ParsePoFileTests(SimpleTestCase):
    def test_missing_required_column_raises(self):
        path = _make_po_excel([], omit_column="Ordered Quantity")
        try:
            with self.assertRaises(POParseError):
                parse_po_file(path)
        finally:
            os.remove(path)

    def test_normal_rows_parsed_correctly(self):
        path = _make_po_excel([SAMPLE_ROW])
        try:
            df = parse_po_file(path)
            self.assertEqual(len(df), 1)
            self.assertEqual(df.iloc[0]["barcode"], "8859388000025")
            self.assertEqual(df.iloc[0]["po_number"], "PO1")
        finally:
            os.remove(path)

    def test_exact_duplicate_rows_are_deduped(self):
        # แถวเดียวกันเป๊ะ (PO+จุดส่ง+SKU+line_no เหมือนกัน) ซ้ำ 2 ครั้ง -> ควรเหลือแค่ 1
        path = _make_po_excel([SAMPLE_ROW, SAMPLE_ROW])
        try:
            df = parse_po_file(path)
            self.assertEqual(len(df), 1)
        finally:
            os.remove(path)

    def test_distinct_rows_not_removed(self):
        row2 = list(SAMPLE_ROW)
        row2[6] = 2  # line_no ต่างกัน -> ไม่ใช่แถวซ้ำ ไม่ควรถูกตัด
        row2[7] = "8859388000026"
        path = _make_po_excel([SAMPLE_ROW, row2])
        try:
            df = parse_po_file(path)
            self.assertEqual(len(df), 2)
        finally:
            os.remove(path)

    def test_row_with_empty_barcode_is_dropped(self):
        row = list(SAMPLE_ROW)
        row[7] = None  # ไม่มีบาร์โค้ด -> เป็นแถวว่าง/สรุปท้ายไฟล์ ควรถูกตัดทิ้ง
        path = _make_po_excel([row])
        try:
            df = parse_po_file(path)
            self.assertEqual(len(df), 0)
        finally:
            os.remove(path)
