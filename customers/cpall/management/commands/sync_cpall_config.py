"""
management command: sync_cpall_config — โหลด SKU master + location mapping จาก YAML เข้า Postgres

วิธีรัน:
    python manage.py sync_cpall_config

ทำเป็น Django management command แทนการรัน "python -m ..." ตรงๆ เพราะตอนนี้ config_loader.py ใช้
Django ORM แล้ว (Phase 1) ซึ่งต้องมี Django app registry setup ก่อนถึงจะ query ได้ — management
command จัดการ setup ให้อัตโนมัติ เป็นวิธีมาตรฐานของ Django สำหรับสคริปต์แบบนี้
"""
from django.core.management.base import BaseCommand

from customers.cpall.logic.config_loader import load_location_mapping, load_sku_master


class Command(BaseCommand):
    help = "Sync SKU master and location mapping จากไฟล์ YAML เข้า Postgres (upsert, รันซ้ำได้เรื่อยๆ)"

    def handle(self, *args, **options):
        n_sku = load_sku_master()
        n_loc = load_location_mapping()
        self.stdout.write(self.style.SUCCESS(f"เสร็จแล้ว — SKU {n_sku} รายการ, Location {n_loc} รายการ"))
