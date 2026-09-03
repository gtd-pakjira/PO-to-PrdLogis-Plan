"""
dev_reset_po_and_plan_id_sequences.py — [เครื่องมือ DEV เท่านั้น] รีเซ็ตตัวนับ id ของ PO/แผนกลับไป
เริ่มที่ 1 ใหม่ — ตั้งใจแยกไฟล์ออกมาต่างหากจาก clear_po_and_plan_data.py (ไม่ฝังเป็น flag ในคำสั่ง
เดิม) เพื่อไม่ให้ปนกับ workflow ปกติ ลดโอกาสรันผิดที่ผิดเวลาในระบบจริง

*** ทำไมถึงไม่แนะนำให้ใช้ในระบบใช้งานจริง (production) ***
PostgreSQL ไม่ลดเลข id คืนให้อัตโนมัติหลัง DELETE เป็นพฤติกรรมมาตรฐาน (ทุกฐานข้อมูลทำแบบนี้) — id
กระโดด (เช่น 7, 8, 9 แทนที่จะเป็น 1, 2, 3) เป็นเรื่องปกติ ไม่กระทบการทำงานอะไรเลย คำสั่งนี้มีไว้แค่
สะดวกเวลาทดสอบแล้วอยากได้เลขสวยๆ — ถ้ามีที่ไหนอ้างอิง id เก่าไว้ (เช่น ลิงก์ที่เคยแชร์) จะชี้ผิดที่ได้
ถ้ารันตอนมีข้อมูลจริงอยู่แล้ว

ต้องรัน clear_po_and_plan_data ให้ตารางว่างเปล่าก่อนเสมอ ถึงจะรีเซ็ตกลับไปที่ 1 ได้จริง (ถ้ายังมีข้อมูล
เหลืออยู่ จะรีเซ็ตไปที่ id สูงสุดที่มีอยู่ +1 แทน กันชนกับข้อมูลเดิม ไม่ทำให้ id ซ้ำกันเด็ดขาด)

*** ทำไมต้องใช้ admin connection ***
role ที่เว็บใช้งานจริง (app_role) มีสิทธิ์แค่ INSERT/SELECT/UPDATE/DELETE บนตาราง ไม่มีสิทธิ์แก้ sequence
โดยตรง (ทดสอบแล้วเจอ "permission denied for sequence" จริงถ้าใช้ Django's connection ปกติ) — ต้องใช้
core.db.get_admin_connection() (superuser) เหมือนตอนรัน schema.sql
"""
from django.core.management.base import BaseCommand

from core.db import get_admin_connection

TABLES = ["po_import", "po_line", "plan_run", "plan_run_logistic_file", "plan_sku_result"]


class Command(BaseCommand):
    help = "[DEV เท่านั้น] รีเซ็ตตัวนับ id ของ PO/แผนกลับไปเริ่มที่ 1 (ปลอดภัยที่สุดเมื่อตารางว่างเปล่าแล้ว)"

    def add_arguments(self, parser):
        parser.add_argument("--yes", action="store_true", help="ข้ามการถามยืนยัน")

    def handle(self, *args, **options):
        self.stdout.write(self.style.WARNING(
            "คำสั่งนี้เป็นเครื่องมือ DEV เท่านั้น ไม่ควรใช้ในระบบที่มีข้อมูลจริงอยู่แล้ว"
        ))
        self.stdout.write("จะรีเซ็ตตัวนับ id ของตาราง: " + ", ".join(TABLES))

        if not options["yes"]:
            answer = input("\nยืนยันรีเซ็ตตัวนับ id ทั้งหมดนี้? พิมพ์ 'yes' เพื่อยืนยัน: ")
            if answer.strip().lower() != "yes":
                self.stdout.write(self.style.WARNING("ยกเลิก ไม่ได้ทำอะไรเลย"))
                return

        conn = get_admin_connection()
        try:
            with conn.cursor() as cursor:
                for table in TABLES:
                    cursor.execute(f"SELECT MAX(id) FROM {table}")
                    max_id = cursor.fetchone()[0]
                    next_value = 1 if max_id is None else max_id + 1
                    cursor.execute(
                        f"SELECT setval(pg_get_serial_sequence('{table}', 'id'), "
                        f"COALESCE((SELECT MAX(id) FROM {table}), 1), "
                        f"(SELECT MAX(id) FROM {table}) IS NOT NULL)"
                    )
                    self.stdout.write(f"  - {table}: ตัวถัดไปจะเริ่มที่ {next_value}")
            conn.commit()
        finally:
            conn.close()

        self.stdout.write(self.style.SUCCESS("รีเซ็ตตัวนับ id เรียบร้อยแล้ว"))

