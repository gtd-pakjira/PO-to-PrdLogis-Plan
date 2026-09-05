"""
date_utils.py — แปลงวันที่ที่ Admin กรอก (ปีค.ศ. ปกติ) เป็นรูปแบบที่เทมเพลตใช้จริง
(วันที่/เดือน/ปีพ.ศ. 2 หลัก เช่น "28/08/69" สำหรับ 28 ส.ค. 2569 = 28 ส.ค. 2026)
แล้วแทนที่ข้อความ "วันที่ ... " ในหัวไฟล์ให้อัตโนมัติ โดยคงข้อความรอบๆ (เช่น "ส่ง PO", "PO.") ไว้เหมือนเดิม
"""
import re
from datetime import date, datetime

DATE_IN_CELL_RE = re.compile(r"\d{1,2}/\d{1,2}/\d{2,4}")


def parse_date_arg(s: str) -> date:
    """รับวันที่จาก Admin ได้หลายรูปแบบ: YYYY-MM-DD, DD/MM/YYYY, DD-MM-YYYY (ปีค.ศ. ปกติ)"""
    s = s.strip()
    for fmt in ("%Y-%m-%d", "%d/%m/%Y", "%d-%m-%Y"):
        try:
            return datetime.strptime(s, fmt).date()
        except ValueError:
            continue
    raise ValueError(f"อ่านวันที่ '{s}' ไม่ออก ใช้รูปแบบ YYYY-MM-DD หรือ DD/MM/YYYY เช่น 2026-08-28")


def format_thai_short(d: date) -> str:
    """แปลงเป็นรูปแบบที่เทมเพลตใช้: DD/MM/YY (ปีพ.ศ. 2 หลัก) เช่น 2026-08-28 -> '28/08/69'"""
    be_year = d.year + 543
    return f"{d.day:02d}/{d.month:02d}/{be_year % 100:02d}"


def fixed_date_resolver(production_date: date, po_date: date):
    """resolver แบบง่าย ใช้วันที่ชุดเดียวกันทั้งไฟล์ (ใช้กับ Logistic Plan ที่แต่ละไฟล์มาจากรอบเดียว)"""
    return lambda col: (production_date, po_date)


def update_date_headers(ws, date_resolver, search_rows=range(1, 15), search_cols=range(1, 30)):
    """
    สแกนหาเซลล์ที่มีคำว่า 'วันที่' แล้วแทนที่ตัวเลขวันที่ในเซลล์นั้นด้วยค่าใหม่
    กติกา: วันที่ตัวแรกที่เจอในเซลล์ = วันที่ผลิต (production_date), ตัวถัดไป = วันที่ PO (po_date)
    (ตรงกับรูปแบบที่เจอจริงในไฟล์บางบัวทอง: "วันที่ผลิต {ผลิต} ส่งวันที่ PO {PO}")
    คงข้อความอื่นในเซลล์ไว้เหมือนเดิมทั้งหมด (ไม่ทับทั้งเซลล์)

    date_resolver: ฟังก์ชัน (col: int) -> (production_date, po_date) | None
    รับเป็นฟังก์ชันแทนที่จะเป็นวันที่ตายตัว เพราะไฟล์ที่รวมหลายรอบ PO เข้าด้วยกัน (เช่น Production Plan)
    อาจมีวันที่หัวไฟล์มากกว่า 1 จุด แต่ละจุดเป็นของคนละรอบ (คนละคอลัมน์) — ถ้า resolver คืน None
    สำหรับคอลัมน์นั้น จะข้ามไม่แก้เซลล์นั้น (กันเผื่อหาไม่ได้ว่าเซลล์นี้เป็นของรอบไหน)

    คืนค่าจำนวนเซลล์ที่แก้ไป
    """
    updated = 0

    for row in search_rows:
        for col in search_cols:
            cell = ws.cell(row=row, column=col)
            val = cell.value
            if not val or "วันที่" not in str(val):
                continue

            dates = date_resolver(col)
            if dates is None:
                continue
            production_date, po_date = dates
            po_str = format_thai_short(po_date)
            production_str = format_thai_short(production_date)

            text = str(val)
            matches = list(DATE_IN_CELL_RE.finditer(text))
            if not matches:
                continue

            replacements = [production_str] + [po_str] * (len(matches) - 1)
            new_text = text
            # แทนที่จากท้ายไปหน้า กันปัญหาตำแหน่ง index เลื่อนตอนความยาวข้อความเปลี่ยน
            for m, repl in reversed(list(zip(matches, replacements))):
                new_text = new_text[:m.start()] + repl + new_text[m.end():]

            cell.value = new_text
            updated += 1

    return updated


def find_merged_date_header_column(ws, row: int = 5, col_filter=None) -> int | None:
    """
    หาคอลัมน์ anchor ของ merged cell ที่ครอบคลุม row ที่กำหนด (default 5) และมีข้อความ "วันที่" อยู่
    — ใช้แก้ปัญหา M5 (หัวไฟล์หลัก "วันที่ผลิต...ส่งวันที่ PO...") ที่เดิม hardcode เป็น col==13 ตรงๆ
    เพราะ M5 บังเอิญเป็น merged cell (M5:Q6) ที่ทับกับคอลัมน์ของจุดส่งย่อยตัวหนึ่งพอดี (2025-09-05) —
    ถ้า Admin แทรกคอลัมน์ใหม่ก่อนคอลัมน์ M ในเทมเพลตทีหลัง merged cell จะเลื่อนไปคอลัมน์อื่น แต่โค้ด
    เดิมยังหา col==13 อยู่ ทำให้วันที่หัวไฟล์ผิดเงียบๆ อีกครั้ง — หาแบบ dynamic จาก merged_cells ของ
    ไฟล์เองแทน ไม่ hardcode เลขคอลัมน์เลย

    *** สำคัญ: เทมเพลต Production Plan มี merged "วันที่" ที่ row นี้ "มากกว่า 1 จุด" จริง ***
    (พิสูจน์แล้วจากการทดสอบ — G5:L6 ของรอบบ่าย กับ M5:Q6 ของรอบเช้าต่างจังหวัด คนละอันกัน) เพราะงั้น
    ฟังก์ชันนี้จะคืนค่า "ตัวแรกที่เจอ" เท่านั้น ถ้าต้องการเจาะจงว่าเอาอันที่ column อยู่ในช่วงไหน
    ให้ส่ง col_filter (function: col:int -> bool) มากรองก่อนคืนค่า — caller ต้องรู้ว่าต้องการอันไหน
    เอง อย่าสมมติว่ามีแค่จุดเดียวเด็ดขาด (เจอบั๊กจริงจากการเขียนฟังก์ชันนี้ครั้งแรกที่ไม่กรอง — คืนค่า
    merged cell ตัวแรกที่เจอ (G5:L6) แทนที่จะเป็น M5:Q6 ที่ต้องการจริง)

    คืนค่า None ถ้าหาไม่เจอ (เช่น เทมเพลตไม่มี merged cell แบบนี้ที่ row นั้นเลย หรือกรองแล้วไม่เจอเลย)
    """
    for merged_range in ws.merged_cells.ranges:
        if not (merged_range.min_row <= row <= merged_range.max_row):
            continue
        if merged_range.min_col == merged_range.max_col:
            continue  # merged cell กว้างแค่คอลัมน์เดียว ไม่ใช่ header หลักแบบที่ต้องการ
        if col_filter is not None and not col_filter(merged_range.min_col):
            continue
        anchor_value = ws.cell(row=merged_range.min_row, column=merged_range.min_col).value
        if anchor_value and "วันที่" in str(anchor_value):
            return merged_range.min_col
    return None
