"""
clear_po_and_plan_data.py — ลบข้อมูล PO/แผนทั้งหมดทิ้ง เพื่อเริ่มทดสอบใหม่ตั้งแต่ศูนย์

ใช้ตอนอยากเคลียร์ข้อมูลทดสอบทั้งหมดก่อนเริ่มใช้งานจริง (หรือทดสอบรอบใหม่) — ไม่แตะ SKU Master,
Location Mapping, Template (เวอร์ชัน/ไฟล์), Customer, กลุ่มพื้นที่ เลย (ข้อมูลตั้งต้นเหล่านี้ยังต้องใช้)

ลบตามลำดับ FK ที่ถูกต้อง (ลูกก่อนแม่) แม้ตารางส่วนใหญ่จะตั้ง CASCADE ไว้อยู่แล้วก็ตาม เพื่อความชัดเจน
และป้องกัน error ถ้ามีคนแก้ CASCADE settings ทีหลัง

ไม่ลบไฟล์บนดิสก์ (customers/cpall/data/output/, data/po_uploads/) ให้อัตโนมัติ — PO/แผนใหม่ทั้งหมด
เป็น data-first อยู่แล้ว (ไม่มีไฟล์ค้าง) มีแค่ PO/แผนเก่าก่อนอัปเดตนี้เท่านั้นที่อาจมีไฟล์ค้างอยู่จริง —
ถ้าอยากลบไฟล์เก่าที่ค้างด้วย ลบโฟลเดอร์เหล่านั้นเองแยกต่างหาก (ไม่กระทบโครงสร้างโฟลเดอร์)
"""
from django.core.management.base import BaseCommand

from customers.cpall.models import (
    PlanRun,
    PlanRunImport,
    PlanRunLogisticFile,
    PlanSkuResult,
    PoImport,
    PoLine,
)


class Command(BaseCommand):
    help = "ลบข้อมูล PO/แผนทั้งหมด (ไม่แตะ SKU/Location/Template/Customer/กลุ่มพื้นที่) — ต้องยืนยันก่อนลบจริง"

    def add_arguments(self, parser):
        parser.add_argument(
            "--yes", action="store_true",
            help="ข้ามการถามยืนยัน (ใช้ตอนรันผ่านสคริปต์อัตโนมัติ) — ปกติควรปล่อยว่างไว้ให้ถามก่อนเสมอ",
        )

    def handle(self, *args, **options):
        counts = {
            "PlanSkuResult": PlanSkuResult.objects.count(),
            "PlanRunLogisticFile": PlanRunLogisticFile.objects.count(),
            "PlanRunImport": PlanRunImport.objects.count(),
            "PlanRun": PlanRun.objects.count(),
            "PoLine": PoLine.objects.count(),
            "PoImport": PoImport.objects.count(),
        }
        self.stdout.write("จะลบข้อมูลต่อไปนี้ทั้งหมด:")
        for name, count in counts.items():
            self.stdout.write(f"  - {name}: {count} แถว")

        if not options["yes"]:
            answer = input("\nยืนยันลบทั้งหมดนี้ถาวร (กู้คืนไม่ได้)? พิมพ์ 'yes' เพื่อยืนยัน: ")
            if answer.strip().lower() != "yes":
                self.stdout.write(self.style.WARNING("ยกเลิก ไม่ได้ลบอะไรเลย"))
                return

        # ลบตามลำดับลูกก่อนแม่ (แม้ตารางส่วนใหญ่ตั้ง CASCADE ไว้แล้ว ทำชัดเจนไว้กันพลาด)
        PlanSkuResult.objects.all().delete()
        PlanRunLogisticFile.objects.all().delete()
        PlanRunImport.objects.all().delete()
        PlanRun.objects.all().delete()
        PoLine.objects.all().delete()
        PoImport.objects.all().delete()

        self.stdout.write(self.style.SUCCESS("ลบข้อมูล PO/แผนทั้งหมดเรียบร้อยแล้ว"))
        self.stdout.write(
            "หมายเหตุ: ไฟล์เก่าที่อาจค้างอยู่บนดิสก์ (PO/แผนก่อนอัปเดต data-first) ไม่ได้ถูกลบให้อัตโนมัติ "
            "— ลบเองได้ที่ customers/cpall/data/output/ และ customers/cpall/data/po_uploads/ ถ้าต้องการ"
        )
