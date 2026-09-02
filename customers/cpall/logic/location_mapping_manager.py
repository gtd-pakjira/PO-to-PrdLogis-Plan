"""
location_mapping_manager.py — ให้ Admin เลือก map รหัสสถานที่ (FC code) ที่ยังไม่รู้จักเข้ากลุ่มพื้นที่
ผ่านหน้าเว็บ (UC-2) แทนที่ต้องไปแก้ไฟล์ customers/cpall/config/location_mapping.yaml เอง

บันทึกทั้ง 2 ที่เสมอ:
  1. ตาราง location_mapping ใน Postgres (ให้ใช้ได้ทันทีในรอบถัดไป) — ผ่าน Django ORM (Phase 1)
  2. ไฟล์ customers/cpall/config/location_mapping.yaml (ให้ค่ายังอยู่ถาวร แม้มีคนรัน config_loader
     ใหม่ทีหลัง ซึ่งจะ sync จาก YAML ทับ DB อีกที — ถ้าไม่เขียนกลับ YAML ด้วย ค่าที่เพิ่มผ่านเว็บจะหาย
     ไปตอน sync รอบหน้า)
"""
import os

from customers.cpall.logic.db import get_cpall_customer_id
from customers.cpall.models import LocationMapping

YAML_PATH = "customers/cpall/config/location_mapping.yaml"


def get_existing_groups() -> list[str]:
    """ดึงชื่อกลุ่มพื้นที่ที่มีอยู่แล้วทั้งหมด (ไว้ให้เลือกในหน้าเว็บ แทนที่จะพิมพ์เอง)"""
    return list(
        LocationMapping.objects.order_by("group").values_list("group", flat=True).distinct()
    )


def save_location_mapping(fc_code: str, name_th: str, group: str, sub_location: str):
    """บันทึก mapping ใหม่ 1 รายการ ทั้งใน Postgres และไฟล์ YAML (ดูเหตุผลใน docstring บนสุดของไฟล์)"""
    customer_id = get_cpall_customer_id()
    LocationMapping.objects.update_or_create(
        fc_code=fc_code,
        defaults={
            "customer_id": customer_id, "name_th": name_th, "group": group, "sub_location": sub_location,
        },
    )
    _append_to_yaml(fc_code, name_th, group, sub_location)


def _append_to_yaml(fc_code: str, name_th: str, group: str, sub_location: str):
    """
    ต่อท้ายไฟล์ YAML แบบ raw text (ไม่ใช้ yaml.dump ทับทั้งไฟล์) เพื่อไม่ทำลาย comment/หมายเหตุ
    ที่มีอยู่แล้วในไฟล์ — เพิ่มเข้าไปแบบเดียวกับ entry อื่นๆ ที่มีอยู่
    """

    def esc(s):
        return str(s).replace('"', '\\"')

    entry = (
        f'\n  - fc_code: "{esc(fc_code)}"\n'
        f'    name_th: "{esc(name_th)}"\n'
        f'    group: "{esc(group)}"\n'
        f'    sub_location: "{esc(sub_location)}"\n'
    )

    if not os.path.exists(YAML_PATH):
        # ไม่ควรเกิดในทางปฏิบัติ (ไฟล์นี้มีอยู่แล้วเสมอ) แต่กันไว้เผื่อ
        with open(YAML_PATH, "w", encoding="utf-8") as f:
            f.write("locations:\n" + entry)
        return

    with open(YAML_PATH, "a", encoding="utf-8") as f:
        f.write(entry)
