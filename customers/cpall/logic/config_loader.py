"""
config_loader.py — โหลด config จาก YAML (sku_master.yaml, location_mapping.yaml)
เข้า Postgres แบบ upsert (แก้ YAML แล้วรันซ้ำได้เรื่อยๆ ไม่ต้องล้างตารางก่อน)

*** ใช้ Django ORM แล้ว (Phase 1) ต้อง setup Django ก่อนถึงจะเรียกได้ — วิธีรันที่ถูกต้อง: ***
    python manage.py sync_cpall_config
(ดู customers/cpall/management/commands/sync_cpall_config.py — ห้ามรันไฟล์นี้ตรงๆ ด้วย
"python -m ..." อีกต่อไป เพราะจะ error "Apps aren't loaded yet" — Django ORM ต้องผ่าน manage.py เท่านั้น)
"""
import yaml

from customers.cpall.logic.db import get_cpall_customer_id
from customers.cpall.models import LocationMapping, SkuMaster


def load_sku_master(path: str = "customers/cpall/config/sku_master.yaml"):
    with open(path, "r", encoding="utf-8") as f:
        data = yaml.safe_load(f)

    customer_id = get_cpall_customer_id()
    skus = data.get("skus", [])
    for sku in skus:
        SkuMaster.objects.update_or_create(
            customer_id=customer_id, barcode=sku["barcode"],
            defaults={
                "name_th": sku["name_th"],
                "name_en": sku.get("name_en"),
                "pack_size": sku["pack_size"],
                "unit_price": sku.get("unit_price"),
                "note": sku.get("note"),
            },
        )
    print(f"[config_loader] synced {len(skus)} SKUs from {path}")
    return len(skus)


def load_location_mapping(path: str = "customers/cpall/config/location_mapping.yaml"):
    with open(path, "r", encoding="utf-8") as f:
        data = yaml.safe_load(f)

    customer_id = get_cpall_customer_id()
    locations = data.get("locations", [])
    for loc in locations:
        LocationMapping.objects.update_or_create(
            customer_id=customer_id, fc_code=loc["fc_code"],
            defaults={
                "name_th": loc["name_th"],
                "group": loc["group"],
                "sub_location": loc.get("sub_location"),
            },
        )
    print(f"[config_loader] synced {len(locations)} locations from {path}")
    return len(locations)
