"""
product_master_manager.py — ให้ Admin กรอกข้อมูลสินค้า (บาร์โค้ด) ที่ยังไม่รู้จักผ่านหน้าเว็บ แทนที่
ต้องไปแก้ไฟล์ customers/cpall/config/sku_master.yaml เอง — ทำตาม pattern เดียวกับ
location_mapping_manager.py ทุกประการ

บันทึกทั้ง 2 ที่เสมอ:
  1. ตาราง product_master ใน Postgres (ให้ใช้ได้ทันทีในรอบถัดไป) — ผ่าน Django ORM
  2. ไฟล์ customers/cpall/config/sku_master.yaml (ให้ค่ายังอยู่ถาวร แม้มีคนรัน config_loader ใหม่
     ทีหลัง ซึ่งจะ sync จาก YAML ทับ DB อีกที — ถ้าไม่เขียนกลับ YAML ด้วย ค่าที่เพิ่มผ่านเว็บจะหายไป
     ตอน sync รอบหน้า)

หมายเหตุสำคัญ: product_master ไม่มีผลต่อการคำนวณแผนเลย (ดู README) ใช้แค่แสดงชื่อสินค้าที่หน้ากรอก
ยอดเผื่อ — ต่างจาก location_mapping ที่มีผลจริงต่อการจัดกลุ่มพื้นที่ เพราะงั้นขั้นตอนนี้ "แนะนำให้กรอก"
ไม่ใช่ "บังคับกรอกก่อนถึงจะไปต่อได้จริง" (แต่ UI ยังคงคล้ายกันเพื่อความสม่ำเสมอ)
"""
import os

from customers.cpall.logic.db import get_cpall_customer_id
from customers.cpall.models import ProductMaster

YAML_PATH = "customers/cpall/config/sku_master.yaml"


def save_product(barcode: str, name_th: str, name_en: str, pack_size: int, unit_price: float | None):
    """บันทึกสินค้าใหม่ 1 รายการ ทั้งใน Postgres และไฟล์ YAML (ดูเหตุผลใน docstring บนสุดของไฟล์)"""
    customer_id = get_cpall_customer_id()
    ProductMaster.objects.update_or_create(
        customer_id=customer_id, barcode=barcode,
        defaults={"name_th": name_th, "name_en": name_en, "pack_size": pack_size, "unit_price": unit_price},
    )
    _append_to_yaml(barcode, name_th, name_en, pack_size, unit_price)


def _append_to_yaml(barcode: str, name_th: str, name_en: str, pack_size: int, unit_price: float | None):
    """ต่อท้ายไฟล์ YAML แบบ raw text (ไม่ใช้ yaml.dump ทับทั้งไฟล์) เพื่อไม่ทำลาย comment ที่มีอยู่แล้ว"""

    def esc(s):
        return str(s).replace('"', '\\"')

    lines = [f'\n  - barcode: "{esc(barcode)}"', f'    name_th: "{esc(name_th)}"']
    if name_en:
        lines.append(f'    name_en: "{esc(name_en)}"')
    lines.append(f"    pack_size: {int(pack_size)}")
    if unit_price is not None:
        lines.append(f"    unit_price: {unit_price}")
    entry = "\n".join(lines) + "\n"

    if not os.path.exists(YAML_PATH):
        # ไม่ควรเกิดในทางปฏิบัติ (ไฟล์นี้มีอยู่แล้วเสมอ) แต่กันไว้เผื่อ
        with open(YAML_PATH, "w", encoding="utf-8") as f:
            f.write("skus:\n" + entry)
        return

    with open(YAML_PATH, "a", encoding="utf-8") as f:
        f.write(entry)
