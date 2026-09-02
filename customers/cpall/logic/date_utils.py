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
