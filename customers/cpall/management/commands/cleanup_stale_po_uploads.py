"""
cleanup_stale_po_uploads.py — ลบไฟล์ที่ค้างอยู่ใน customers/cpall/data/po_uploads/ ทิ้ง

ระบบเป็น data-first เต็มรูปแบบ — ไฟล์ที่ import สำเร็จ (ไม่ว่าจะผ่านหรือไม่ผ่าน duplicate-confirm)
ถูกลบทิ้งทันทีหลัง parse เสร็จเสมอ เพราะงั้นไฟล์ที่ยัง "ค้าง" อยู่ในโฟลเดอร์นี้มีสาเหตุเดียวที่เป็นไปได้
ในสภาวะใช้งานปกติ: Admin เจอหน้า "พบรายการที่อาจซ้ำ" (confirm_duplicates) แล้วปิด browser หนีไปเฉยๆ
โดยไม่กดปุ่ม "ดำเนินการต่อ" หรือ "ยกเลิก" เลย — ไฟล์ที่ session เก็บ path ไว้เลยไม่มีใครมาลบให้

รันเป็น cron/scheduled task ได้ (เช่น รายวัน) — ลบเฉพาะไฟล์ที่เก่ากว่า --hours ชั่วโมง (default 24)
เท่านั้น กันลบไฟล์ที่เพิ่ง upload ไปหมาดๆ ระหว่างที่ Admin ยังเปิดหน้า confirm ค้างอยู่จริงๆ
"""
import os
import time

from django.core.management.base import BaseCommand

UPLOAD_DIR = "customers/cpall/data/po_uploads"


class Command(BaseCommand):
    help = "ลบไฟล์ PO ที่ค้างอยู่ใน po_uploads/ ทิ้ง (ไฟล์ที่ Admin ไม่เคยกด Continue/Stop ที่หน้ายืนยัน duplicate)"

    def add_arguments(self, parser):
        parser.add_argument(
            "--hours", type=int, default=24,
            help="ลบเฉพาะไฟล์ที่เก่ากว่ากี่ชั่วโมง (default: 24) — กันลบไฟล์ที่เพิ่งค้างระหว่าง Admin ยังเปิดหน้าอยู่จริง",
        )
        parser.add_argument(
            "--dry-run", action="store_true",
            help="แสดงรายชื่อไฟล์ที่จะลบเฉยๆ ไม่ลบจริง",
        )

    def handle(self, *args, **options):
        if not os.path.isdir(UPLOAD_DIR):
            self.stdout.write(self.style.WARNING(f"ไม่พบโฟลเดอร์ {UPLOAD_DIR} เลย ไม่มีอะไรต้องทำ"))
            return

        cutoff = time.time() - options["hours"] * 3600
        stale_files = []
        for name in os.listdir(UPLOAD_DIR):
            if name.startswith("."):
                continue  # ข้าม .gitkeep และไฟล์ระบบอื่นๆ ที่ขึ้นต้นด้วยจุด — เจอบั๊กจริงตอนทดสอบ
                # (ลบ .gitkeep ไปด้วยเพราะมันก็ "เก่ากว่า cutoff" เหมือนกัน ทั้งที่ไม่ควรลบเลย)
            path = os.path.join(UPLOAD_DIR, name)
            if os.path.isfile(path) and os.path.getmtime(path) < cutoff:
                stale_files.append(path)

        if not stale_files:
            self.stdout.write(self.style.SUCCESS(f"ไม่พบไฟล์ค้างที่เก่ากว่า {options['hours']} ชั่วโมงเลย"))
            return

        self.stdout.write(f"พบไฟล์ค้าง {len(stale_files)} ไฟล์ (เก่ากว่า {options['hours']} ชั่วโมง):")
        for path in stale_files:
            self.stdout.write(f"  - {path}")

        if options["dry_run"]:
            self.stdout.write(self.style.WARNING("(--dry-run: ไม่ได้ลบจริง)"))
            return

        for path in stale_files:
            os.remove(path)
        self.stdout.write(self.style.SUCCESS(f"ลบไฟล์ค้างทิ้งเรียบร้อย {len(stale_files)} ไฟล์"))
