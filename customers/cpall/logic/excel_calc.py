"""
excel_calc.py — คำนวณสูตรใน .xlsx จริงด้วย LibreOffice (headless) แล้วอ่านค่าที่คำนวณแล้วกลับมา

*** ทำไมต้องมีไฟล์นี้ ***
openpyxl เขียนสูตรลงไฟล์ได้ แต่ "อ่านค่าที่คำนวณจากสูตรไม่ได้" (ไฟล์ที่ openpyxl สร้างเองไม่เคยผ่าน
โปรแกรม spreadsheet จริงมาก่อน จึงยังไม่มีค่าที่คำนวณเก็บไว้ในไฟล์เลย) — เดิมเราแก้ปัญหานี้ด้วยการ
"จำลองสูตร Excel เป็นโค้ด Python" (ดู plan_view_data.py: _pack_breakdown_text, _basket_total, ยอดคืน)
ซึ่งเสี่ยงตรงที่ ถ้า Admin แก้สูตรในไฟล์เทมเพลตจริง โค้ดเราจะไม่รู้เลย และจะให้ตัวเลขผิดแบบเงียบๆ

ไฟล์นี้ตัดปัญหานั้นที่ราก: ให้ LibreOffice (โปรแกรม spreadsheet จริง) เปิดไฟล์ คำนวณสูตรทั้งหมด แล้ว
เซฟใหม่ จากนั้นอ่านค่าที่คำนวณแล้วด้วย openpyxl (data_only=True) — สูตรในเทมเพลตเปลี่ยนยังไง ค่าที่ได้
ก็เปลี่ยนตามเสมอ ไม่ต้องแก้โค้ด Python ตาม

*** ผลการทดสอบจริงในโปรเจกต์นี้ (2026-08-31) ***
- ค่าที่คำนวณตรงกับที่เคยยืนยันกับไฟล์จริงของ Admin ทุกตัว: ยอดรวม 975, ยอดเผื่อ 260, ยอดคืน 45,
  ตะกร้า 13 (สูตร SUMPRODUCT+ROUNDUP ข้ามชีตก็คำนวณถูก)
- ใช้เวลา ~1 วินาทีต่อไฟล์ รันซ้ำได้ผลเท่ากันทุกครั้ง
- รันพร้อมกันหลาย process ได้ ถ้าแยก UserInstallation profile ต่อ process (สำคัญ — ถ้าใช้ profile
  ร่วมกัน LibreOffice จะชนกันเอง) ฟังก์ชันนี้สร้าง profile ชั่วคราวใหม่ทุกครั้งจึงปลอดภัยโดยธรรมชาติ
"""
import os
import shutil
import subprocess
import tempfile

import openpyxl


class ExcelCalcError(Exception):
    pass


def recalculate(xlsx_path: str, timeout_sec: int = 120) -> str:
    """
    ให้ LibreOffice เปิดไฟล์ คำนวณสูตรทั้งหมด แล้วเซฟเป็นไฟล์ใหม่ คืน path ของไฟล์ผลลัพธ์
    (ไฟล์ต้นฉบับไม่ถูกแก้ไข) — ผู้เรียกมีหน้าที่ลบไฟล์ผลลัพธ์ทิ้งเองเมื่อใช้เสร็จ
    """
    if not os.path.exists(xlsx_path):
        raise ExcelCalcError(f"ไม่พบไฟล์ {xlsx_path}")

    out_dir = tempfile.mkdtemp(prefix="xlsxcalc_out_")
    # profile แยกต่อการเรียก 1 ครั้ง — กัน LibreOffice หลาย process ชนกัน (ทดสอบยืนยันแล้วว่าจำเป็น)
    profile_dir = tempfile.mkdtemp(prefix="xlsxcalc_profile_")

    try:
        result = subprocess.run(
            [
                "soffice", "--headless", "--norestore",
                f"-env:UserInstallation=file://{profile_dir}",
                "--convert-to", "xlsx", "--outdir", out_dir, xlsx_path,
            ],
            capture_output=True, text=True, timeout=timeout_sec,
        )
        if result.returncode != 0:
            raise ExcelCalcError(f"LibreOffice ล้มเหลว: {result.stderr.strip() or result.stdout.strip()}")

        out_path = os.path.join(out_dir, os.path.basename(xlsx_path))
        if not os.path.exists(out_path):
            raise ExcelCalcError(f"LibreOffice ไม่ได้สร้างไฟล์ผลลัพธ์ที่ {out_path}")
        return out_path

    except subprocess.TimeoutExpired:
        shutil.rmtree(out_dir, ignore_errors=True)
        raise ExcelCalcError(f"LibreOffice ใช้เวลานานเกิน {timeout_sec} วินาที")
    except Exception:
        shutil.rmtree(out_dir, ignore_errors=True)
        raise
    finally:
        shutil.rmtree(profile_dir, ignore_errors=True)


def load_calculated_workbook(xlsx_path: str):
    """
    คำนวณไฟล์แล้วคืน openpyxl workbook ที่อ่านค่าที่คำนวณแล้วได้ (data_only=True)
    จัดการลบไฟล์ชั่วคราวให้เอง — ใช้ตัวนี้แทน recalculate() ถ้าไม่ต้องการไฟล์ผลลัพธ์เก็บไว้
    """
    calc_path = recalculate(xlsx_path)
    calc_dir = os.path.dirname(calc_path)
    try:
        wb = openpyxl.load_workbook(calc_path, data_only=True)
        # โหลดเข้า memory แล้ว ลบไฟล์ได้เลย (openpyxl อ่านทั้งไฟล์ตอน load ไม่ได้ lazy-read ทีหลัง)
        return wb
    finally:
        shutil.rmtree(calc_dir, ignore_errors=True)
