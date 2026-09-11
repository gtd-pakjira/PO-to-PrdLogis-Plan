"""
template_manager.py — ดาวน์โหลด/อัปโหลด/ตรวจสอบไฟล์ Template (Production Plan + Logistic Plan 4 กลุ่ม)
พร้อมระบบ versioning เต็มรูปแบบ (Phase 1.6 sub-phase 1)

Admin แก้ไฟล์ Template เองบ่อย (เพิ่มคอลัมน์ PO, เพิ่มแถว SKU ใหม่, แก้สูตร) — โมดูลนี้:
  1. ตรวจสอบโครงสร้างไฟล์ใหม่ทันทีตอนอัปโหลด โดยใช้ตัวสแกนเดียวกับที่ excel_export.py /
     logistic_plan_export.py ใช้ตอนสร้างแผนจริง — ถ้าไฟล์ใหม่ผ่านตรงนี้ แปลว่าใช้สร้างแผนได้แน่นอน
  2. เก็บทุกเวอร์ชันที่เคยอัปโหลดไว้ถาวร (ไม่ใช่แค่ backup 1 ชั้นแบบเดิม) — กู้คืนไปเวอร์ชันไหนก็ได้
     ในประวัติ ลบเวอร์ชันเก่าได้ (ถ้าไม่ใช่เวอร์ชันสุดท้ายที่เหลือ และไม่มีแผนไหนอ้างอิงอยู่)

*** สถาปัตยกรรม: ยังมี "ไฟล์ live" ที่ path ตายตัวเดิม (get_template_registry()[key]["path"]) ***
excel_export.py / logistic_plan_export.py อ่านจาก path ตายตัวนี้เสมอ (ไม่ได้แก้ให้รู้จัก versioning
โดยตรง) — ทุกครั้งที่เปลี่ยนเวอร์ชัน active (อัปโหลดใหม่/กู้คืน) จะ sync ไฟล์ที่ path ตายตัวนี้ให้ตรงกับ
เวอร์ชัน active เสมอ วิธีนี้ทำให้โค้ดเดิมที่มีอยู่แล้วทำงานถูกต้องต่อไปโดยไม่ต้องแก้เลย
"""
import os
import shutil

import openpyxl
from openpyxl.utils import get_column_letter

from customers.cpall.logic.db import get_cpall_customer_id
from customers.cpall.logic.excel_export import TEMPLATE_PATH as PP_TEMPLATE_PATH
from customers.cpall.logic.excel_export import _find_sku_header_rows as _find_pp_sku_header_rows
from customers.cpall.logic.excel_export import _find_sub_location_columns, _find_total_column
from customers.cpall.logic.excel_export import get_sheet_name as get_pp_sheet_name
from customers.cpall.logic.logistic_plan_export import (
    _find_line_no_column,
    _find_qty_column_range,
    get_group_templates,
)
from customers.cpall.logic.logistic_plan_export import _find_sku_header_rows as _find_lp_sku_header_rows

VERSIONS_DIR = "customers/cpall/excel_templates/versions"


class TemplateValidationError(Exception):
    pass


class TemplateInUseError(Exception):
    """ลบเวอร์ชันนี้ไม่ได้ เพราะเป็นเวอร์ชันสุดท้ายที่เหลือ หรือมีแผนอ้างอิงอยู่"""
    pass


def get_template_registry() -> dict:
    """
    สร้าง registry ของ Template ทั้งหมด (Production Plan + Logistic Plan ทุกกลุ่ม) — query กลุ่มพื้นที่
    สดจาก database ทุกครั้ง (ผ่าน get_group_templates()) แทนที่จะ build ครั้งเดียวตอน import module
    เหมือนเดิม (TEMPLATE_REGISTRY เคยเป็น module-level constant) — เพิ่ม/ปิดใช้งานกลุ่มพื้นที่ผ่าน
    หน้าเว็บแล้วเห็นผลทันที ไม่ต้อง restart เซิร์ฟเวอร์
    """
    registry = {
        "production_plan": {
            "path": PP_TEMPLATE_PATH,
            "label": "แพลนผลิต",
            "kind": "production",
        }
    }
    for group_name, (path, _) in get_group_templates().items():
        registry[f"logistic_{group_name}"] = {
            "path": path,
            "label": f"แพลนกระจาย — {group_name}",
            "kind": "logistic",
            "group": group_name,
        }
    return registry


def get_template_grid(key: str, sheet_name: str = None, max_rows: int = 120, max_cols: int = 25) -> dict:
    """
    อ่านไฟล์ Template แบบดิบๆ (ไม่คำนวณ ไม่แปลงอะไรเลย) มาจัดเป็นตารางสำหรับแสดงในเว็บ — เห็นสูตร
    Excel ตรงๆ (เช่น "=SUM(G11:Q11)") แทนที่จะเห็นแค่ค่าที่คำนวณแล้ว — อ่านจากไฟล์ live เสมอ (เวอร์ชัน
    active ปัจจุบัน) ไม่มีการแก้ไขไฟล์ใดๆ ในฟังก์ชันนี้ อ่านอย่างเดียว (read-only)

    sheet_name: ถ้าไม่ระบุ ใช้ชีตหลักของ template นี้ (ตามที่ลงทะเบียนไว้) — ไฟล์เดียวอาจมีหลายชีต
    (เช่น ชีต "-รถ"/"คันที่ 1/2" ที่เกี่ยวกับจัดรถ) เลือกดูชีตอื่นได้ผ่าน sheet_name
    """
    registry = get_template_registry()
    if key not in registry:
        raise TemplateValidationError(f"ไม่รู้จัก template '{key}'")

    info = registry[key]
    path = info["path"]
    if not os.path.exists(path):
        raise TemplateValidationError(f"ไม่พบไฟล์ {path}")

    wb = openpyxl.load_workbook(path, data_only=False)  # data_only=False สำคัญมาก — เอาไว้เห็นสูตรดิบ

    if info["kind"] == "production":
        default_sheet = get_pp_sheet_name()
    else:
        _, default_sheet = get_group_templates()[info["group"]]

    if sheet_name is None or sheet_name not in wb.sheetnames:
        sheet_name = default_sheet if default_sheet in wb.sheetnames else wb.sheetnames[0]

    ws = wb[sheet_name]
    n_rows = min(ws.max_row or 1, max_rows)
    n_cols = min(ws.max_column or 1, max_cols)

    col_letters = [get_column_letter(c) for c in range(1, n_cols + 1)]
    grid_rows = []
    for r in range(1, n_rows + 1):
        cells = []
        for c in range(1, n_cols + 1):
            val = ws.cell(row=r, column=c).value
            is_formula = isinstance(val, str) and val.startswith("=")
            cells.append({"value": val, "is_formula": is_formula})
        grid_rows.append({"row_num": r, "cells": cells})

    return {
        "label": info["label"],
        "sheet_names": wb.sheetnames,
        "current_sheet": sheet_name,
        "col_letters": col_letters,
        "rows": grid_rows,
        "truncated": (ws.max_row or 1) > max_rows or (ws.max_column or 1) > max_cols,
    }


def validate_template(key: str, filepath: str) -> dict:
    """
    ตรวจสอบโครงสร้างไฟล์ Template ที่อัปโหลดมา
    raise TemplateValidationError พร้อมเหตุผลชัดเจนถ้าไม่ผ่าน คืนค่าสรุปข้อมูลถ้าผ่าน (ไว้โชว์ผู้ใช้)
    """
    registry = get_template_registry()
    if key not in registry:
        raise TemplateValidationError(f"ไม่รู้จัก template '{key}'")

    info = registry[key]
    try:
        wb = openpyxl.load_workbook(filepath)
    except Exception as e:
        raise TemplateValidationError(f"เปิดไฟล์ไม่ได้ (ไม่ใช่ไฟล์ .xlsx ที่ถูกต้อง): {e}")

    if info["kind"] == "production":
        pp_sheet_name = get_pp_sheet_name()
        if pp_sheet_name not in wb.sheetnames:
            raise TemplateValidationError(
                f"ไม่พบชีตชื่อ '{pp_sheet_name}' ในไฟล์ — เช็คว่าไม่ได้เปลี่ยนชื่อชีตตอนแก้ไฟล์"
            )
        ws = wb[pp_sheet_name]
        try:
            col_to_sub = _find_sub_location_columns(ws)
            total_col = _find_total_column(ws)
            header_rows = _find_pp_sku_header_rows(ws)
        except Exception as e:
            raise TemplateValidationError(f"อ่านโครงสร้างไฟล์ไม่ผ่าน: {e}")

        if not col_to_sub:
            raise TemplateValidationError("หาคอลัมน์จุดส่งย่อยไม่เจอ (แถวหัวตาราง 7-8 อาจถูกแก้ผิดรูปแบบ)")
        if total_col is None:
            raise TemplateValidationError("หาคอลัมน์ 'ยอดรวม' ไม่เจอ")
        if not header_rows:
            raise TemplateValidationError("หาแถว SKU ไม่เจอเลยสักแถว")

        return {"sub_locations": list(col_to_sub.values()), "sku_count": len(header_rows)}

    else:  # logistic
        group_name = info["group"]
        _, sheet_name = get_group_templates()[group_name]
        if sheet_name not in wb.sheetnames:
            raise TemplateValidationError(
                f"ไม่พบชีตชื่อ '{sheet_name}' ในไฟล์ — เช็คว่าไม่ได้เปลี่ยนชื่อชีตตอนแก้ไฟล์"
            )
        ws = wb[sheet_name]
        try:
            line_no_col, header_row = _find_line_no_column(ws)
            name_col = line_no_col + 1
            qty_start_col = line_no_col + 3
            qty_start_col, qty_end_col = _find_qty_column_range(ws, qty_start_col, header_row)
            header_rows = _find_lp_sku_header_rows(ws, name_col)
        except Exception as e:
            raise TemplateValidationError(f"อ่านโครงสร้างไฟล์ไม่ผ่าน: {e}")

        if not header_rows:
            raise TemplateValidationError("หาแถว SKU ไม่เจอเลยสักแถว")

        return {"qty_column_count": qty_end_col - qty_start_col + 1, "sku_count": len(header_rows)}

def reconcile_template_with_product_master(key: str, filepath: str) -> dict:
    """
    ตรวจสอบความสอดคล้องระหว่าง Template กับ ProductMaster

    Production Template:
      - Template มี + ProductMaster ไม่มี      -> CREATE_PRODUCT
      - Template มี + ProductMaster inactive    -> ACTIVATE
      - ProductMaster active + Template ไม่มี   -> DEACTIVATE
      - ทั้งคู่ active / ทั้งคู่ไม่มี            -> MATCH

    Logistic Template:
      - Template มี + ProductMaster ไม่มี      -> CREATE_PRODUCT
      - Template มี + ProductMaster inactive    -> ACTIVATE
      - ProductMaster active แต่ไม่อยู่ในกลุ่มนี้
        -> ไม่ถือว่า mismatch เพราะ Logistic Template
           เป็น subset ตาม Location/Group
    """
    registry = get_template_registry()

    if key not in registry:
        raise TemplateValidationError(f"ไม่รู้จัก template '{key}'")

    info = registry[key]

    try:
        wb = openpyxl.load_workbook(filepath, data_only=False)
    except Exception as e:
        raise TemplateValidationError(f"เปิดไฟล์ไม่ได้: {e}")

    # ---------------------------------------------------------
    # 1) อ่าน SKU จาก Template
    # ---------------------------------------------------------
    if info["kind"] == "production":
        sheet_name = get_pp_sheet_name()

        if sheet_name not in wb.sheetnames:
            raise TemplateValidationError(
                f"ไม่พบชีต '{sheet_name}' ใน Template"
            )

        ws = wb[sheet_name]
        template_rows = _find_pp_sku_header_rows(ws)

        # Production:
        # header row -> ชื่อสินค้าอยู่ col D
        template_items = {}

        for barcode, row in template_rows.items():
            item_name = str(ws.cell(row=row, column=4).value or "").strip()

            template_items[barcode] = {
                "barcode": barcode,
                "item_name": item_name or None,
                "row": row,
            }

    else:
        # Logistic
        group_name = info["group"]
        _, sheet_name = get_group_templates()[group_name]

        if sheet_name not in wb.sheetnames:
            raise TemplateValidationError(
                f"ไม่พบชีต '{sheet_name}' ใน Template"
            )

        ws = wb[sheet_name]

        line_no_col, header_row = _find_line_no_column(ws)
        name_col = line_no_col + 1

        template_rows = _find_lp_sku_header_rows(ws, name_col)

        template_items = {}

        for barcode, row in template_rows.items():
            item_name = str(ws.cell(row=row, column=name_col).value or "").strip()

            template_items[barcode] = {
                "barcode": barcode,
                "item_name": item_name or None,
                "row": row,
            }

    # ---------------------------------------------------------
    # 2) อ่าน ProductMaster
    # ---------------------------------------------------------
    from customers.cpall.models import ProductMaster

    customer_id = get_cpall_customer_id()

    products = {
        p.barcode: p
        for p in ProductMaster.objects.filter(customer_id=customer_id)
    }

    # ---------------------------------------------------------
    # 3) Reconcile
    # ---------------------------------------------------------
    matched = []
    missing_in_product_master = []
    inactive_in_product_master = []
    active_missing_in_template = []

    # Template -> ProductMaster
    for barcode, template_item in sorted(template_items.items()):

        product = products.get(barcode)

        # Template มี แต่ ProductMaster ไม่มี
        if product is None:
            missing_in_product_master.append({
                "barcode": barcode,
                "item_name": template_item["item_name"],
                "template_row": template_item["row"],
                "action": "CREATE_PRODUCT",
            })
            continue

        # Template มี แต่ ProductMaster inactive
        if not product.is_active:
            inactive_in_product_master.append({
                "barcode": barcode,
                "item_name": product.name_th or template_item["item_name"],
                "template_item_name": template_item["item_name"],
                "template_row": template_item["row"],
                "action": "ACTIVATE",
            })
            continue

        # ตรงกัน
        matched.append({
            "barcode": barcode,
            "item_name": product.name_th or template_item["item_name"],
            "template_row": template_item["row"],
        })

    # ---------------------------------------------------------
    # 4) ProductMaster -> Template
    #
    # เฉพาะ Production Template เท่านั้น
    # Logistic Template เป็น subset ของสินค้า
    # ---------------------------------------------------------
    if info["kind"] == "production":
        template_barcodes = set(template_items.keys())

        for barcode, product in sorted(products.items()):
            if product.is_active and barcode not in template_barcodes:
                active_missing_in_template.append({
                    "barcode": barcode,
                    "item_name": product.name_th,
                    "action": "DEACTIVATE",
                })

    # ---------------------------------------------------------
    # 5) Summary
    # ---------------------------------------------------------
    mismatch_count = (
        len(missing_in_product_master)
        + len(inactive_in_product_master)
        + len(active_missing_in_template)
    )

    return {
        "template_key": key,
        "template_label": info["label"],
        "kind": info["kind"],

        "template_sku_count": len(template_items),
        "product_master_sku_count": len(products),

        "matched": matched,
        "missing_in_product_master": missing_in_product_master,
        "inactive_in_product_master": inactive_in_product_master,
        "active_missing_in_template": active_missing_in_template,

        "matched_count": len(matched),
        "missing_count": len(missing_in_product_master),
        "inactive_count": len(inactive_in_product_master),
        "active_missing_count": len(active_missing_in_template),

        "mismatch_count": mismatch_count,
        "is_consistent": mismatch_count == 0,
    }

def get_template_reconciliation(key: str, filepath: str) -> dict:
    """
    ตรวจสอบ Template กับ ProductMaster และคืนผลสำหรับหน้า Admin

    ไม่แก้ ProductMaster
    ไม่เปลี่ยน active version
    อ่านอย่างเดียว
    """
    result = reconcile_template_with_product_master(key, filepath)

    return {
        "template_key": key,
        "template_label": result["template_label"],
        "kind": result["kind"],

        "is_consistent": result["is_consistent"],
        "mismatch_count": result["mismatch_count"],

        "matched_count": result["matched_count"],
        "missing_count": result["missing_count"],
        "inactive_count": result["inactive_count"],
        "active_missing_count": result["active_missing_count"],

        "matched": result["matched"],
        "missing_in_product_master": result["missing_in_product_master"],
        "inactive_in_product_master": result["inactive_in_product_master"],
        "active_missing_in_template": result["active_missing_in_template"],
    }

def reconcile_template_version(key: str, version_id: int) -> dict:
    """
    Reconcile TemplateVersion ที่ระบุ กับ ProductMaster ปัจจุบัน
    โดยไม่เปลี่ยนสถานะอะไร
    """
    from customers.cpall.models import TemplateVersion

    try:
        version = TemplateVersion.objects.get(
            id=version_id,
            template_key=key,
        )
    except TemplateVersion.DoesNotExist:
        raise TemplateValidationError("ไม่พบเวอร์ชันนี้")

    result = reconcile_template_with_product_master(
        key,
        version.file_path,
    )

    result["version_id"] = version.id
    result["version_number"] = version.version_number

    return result

def build_group_reconcile_actions(all_versions, reconciles):
    """
    รวม ProductMaster actions ของทุก Template ใน Group
    ให้ SKU + action เดียวกันเหลือเพียง 1 action

    ตัวอย่าง:
        Production -> DEACTIVATE SKU-A
        Logistic A -> DEACTIVATE SKU-A
        Logistic B -> DEACTIVATE SKU-A

    จะเหลือ:
        DEACTIVATE SKU-A
        sources = [Production, Logistic A, Logistic B]

    ยังไม่แก้ ProductMaster
    """
    actions = {}
    conflicts = {}

    for version, reconcile in zip(all_versions, reconciles):
        source = reconcile["template_label"]

        action_items = [
            *reconcile["missing_in_product_master"],
            *reconcile["inactive_in_product_master"],
            *reconcile["active_missing_in_template"],
        ]

        for item in action_items:
            action = item["action"]
            barcode = str(item["barcode"]).strip()

            key = (action, barcode)

            if key not in actions:
                actions[key] = {
                    "action": action,
                    "barcode": barcode,
                    "item_name": (
                        item.get("item_name")
                        or item.get("template_item_name")
                        or ""
                    ),
                    "sources": [source],
                }
                continue

            # action + barcode ซ้ำกัน → รวม source
            if source not in actions[key]["sources"]:
                actions[key]["sources"].append(source)

    # ตรวจ conflict:
    # SKU เดียวกัน แต่มีคนละ action เช่น
    # ACTIVATE + SKU-A
    # DEACTIVATE + SKU-A
    by_barcode = {}

    for action_data in actions.values():
        barcode = action_data["barcode"]
        by_barcode.setdefault(barcode, []).append(action_data)

    for barcode, barcode_actions in by_barcode.items():
        action_names = {item["action"] for item in barcode_actions}

        if len(action_names) > 1:
            conflicts[barcode] = barcode_actions

    return {
        "actions": list(actions.values()),
        "conflicts": list(conflicts.values()),
    }


def apply_product_master_action(
    action: str,
    barcode: str,
    *,
    name_th: str = None,
    name_en: str = None,
    pack_size=None,
    unit_price=None,
) -> dict:
    """
    ให้ Admin ยืนยันการเปลี่ยน ProductMaster เป็นราย SKU

    action:
      CREATE_PRODUCT
      ACTIVATE
      DEACTIVATE
    """
    from django.db import transaction
    from customers.cpall.models import ProductMaster
    from customers.cpall.logic.product_master_manager import save_product

    customer_id = get_cpall_customer_id()

    barcode = str(barcode).strip()

    if not barcode:
        raise TemplateValidationError("ไม่พบ Barcode")

    with transaction.atomic():

        if action == "CREATE_PRODUCT":

            if not name_th:
                raise TemplateValidationError(
                    f"SKU {barcode} ต้องระบุชื่อสินค้า"
                )

            product, created = ProductMaster.objects.get_or_create(
                customer_id=customer_id,
                barcode=barcode,
                defaults={
                    "name_th": name_th,
                    "name_en": name_en or "",
                    "pack_size": pack_size,
                    "unit_price": unit_price,
                    "is_active": True,
                },
            )

            if not created:
                raise TemplateValidationError(
                    f"SKU {barcode} มีอยู่ใน ProductMaster แล้ว"
                )

            return {
                "action": action,
                "barcode": barcode,
                "product_name": product.name_th,
                "status": "created",
            }

        product = ProductMaster.objects.filter(
            customer_id=customer_id,
            barcode=barcode,
        ).first()

        if product is None:
            raise TemplateValidationError(
                f"ไม่พบ SKU {barcode} ใน ProductMaster"
            )

        if action == "ACTIVATE":
            product.is_active = True
            product.save(update_fields=["is_active", "updated_at"])

            return {
                "action": action,
                "barcode": barcode,
                "product_name": product.name_th,
                "status": "activated",
            }

        if action == "DEACTIVATE":
            product.is_active = False
            product.save(update_fields=["is_active", "updated_at"])

            return {
                "action": action,
                "barcode": barcode,
                "product_name": product.name_th,
                "status": "deactivated",
            }

        raise TemplateValidationError(
            f"ไม่รู้จัก action '{action}'"
        )

def activate_template_group(group, product_actions=None):
    """
    Activate TemplateGroup พร้อม ProductMaster changes แบบ atomic

    product_actions:
        [
            {
                "action": "CREATE_PRODUCT",
                "barcode": "...",
                "name_th": "...",
                "name_en": "...",
                "pack_size": "...",
                "unit_price": "...",
            },
            ...
        ]

    ยังไม่ sync filesystem จนกว่า DB transaction จะสำเร็จ
    """
    from django.db import transaction
    from django.utils import timezone
    from customers.cpall.models import TemplateGroup, TemplateVersion

    product_actions = product_actions or []

    with transaction.atomic():

        # ----------------------------------------
        # 1. ProductMaster
        # ----------------------------------------
        for item in product_actions:
            apply_product_master_action(
                action=item["action"],
                barcode=item["barcode"],
                name_th=item.get("name_th"),
                name_en=item.get("name_en"),
                pack_size=item.get("pack_size"),
                unit_price=item.get("unit_price"),
            )

        # ----------------------------------------
        # 2. หา TemplateVersion ใน Group
        # ----------------------------------------
        group_items = list(
            group.items.select_related("template_version").all()
        )

        # ----------------------------------------
        # 3. Activate TemplateVersion
        # ----------------------------------------
        versions = []

        for item in group_items:
            version = item.template_version
            versions.append(version)

            TemplateVersion.objects.filter(
                template_key=version.template_key,
                is_active=True,
            ).exclude(id=version.id).update(
                is_active=False
            )

            version.is_active = True
            version.save(update_fields=["is_active"])

        # ----------------------------------------
        # 4. ปิด Group เก่า
        # ----------------------------------------
        TemplateGroup.objects.filter(
            is_active=True
        ).exclude(
            id=group.id
        ).update(
            is_active=False
        )

        # ----------------------------------------
        # 5. Activate Group ใหม่
        # ----------------------------------------
        group.is_active = True
        group.activated_at = timezone.now()
        group.save(
            update_fields=["is_active", "activated_at"]
        )

    # ----------------------------------------
    # 6. Sync live files หลัง DB transaction
    # ----------------------------------------
    for version in versions:
        _sync_live_file(
            version.template_key,
            version,
        )

    return versions

def _version_file_path(key: str, version_number: int) -> str:
    return os.path.join(VERSIONS_DIR, key, f"v{version_number}.xlsx")


def _sync_live_file(key: str, version):
    """คัดลอกไฟล์ของเวอร์ชันที่ระบุไปทับไฟล์ live ที่ path ตายตัว — ให้โค้ดเดิมอ่านถูกเวอร์ชันเสมอ"""
    shutil.copy2(version.file_path, get_template_registry()[key]["path"])

def get_active_group_version(key: str):
    """
    คืน TemplateVersion ที่ถือว่า "ใช้งานจริง" ตาม Active TemplateGroup เท่านั้น
    """
    from customers.cpall.models import TemplateGroupItem

    item = (
        TemplateGroupItem.objects
        .filter(
            template_group__customer_id=get_cpall_customer_id(),
            template_group__is_active=True,
            template_version__template_key=key,
        )
        .select_related("template_version")
        .first()
    )

    return item.template_version if item else None


def _ensure_initial_version(key: str):
    """
    ถ้ายังไม่เคยมี TemplateVersion เลยสำหรับ key นี้ (ระบบเพิ่งอัปเกรดมาใช้ versioning ครั้งแรก)
    ให้สร้างเวอร์ชัน 1 จากไฟล์ live ปัจจุบันให้อัตโนมัติ — ไม่ต้องให้ Admin ทำอะไรเพิ่ม
    """
    from customers.cpall.models import TemplateVersion

    if TemplateVersion.objects.filter(template_key=key).exists():
        return

    target_path = get_template_registry()[key]["path"]
    if not os.path.exists(target_path):
        return  # ไม่มีไฟล์ live เลย ไม่มีอะไรให้สร้างเป็นเวอร์ชันแรก

    version_path = _version_file_path(key, 1)
    os.makedirs(os.path.dirname(version_path), exist_ok=True)
    shutil.copy2(target_path, version_path)

    try:
        validation = validate_template(key, target_path)
        summary = f"sku_count={validation.get('sku_count')}"
    except TemplateValidationError:
        summary = "(ไฟล์เดิมก่อนมีระบบ versioning — ไม่ได้ validate ตอนสร้างเวอร์ชันนี้)"

    TemplateVersion.objects.create(
        customer_id=get_cpall_customer_id(),
        template_key=key,
        version_number=1,
        file_path=version_path,
        is_active=False,
        validation_summary=summary,
    )


def list_templates() -> list[dict]:
    """คืนรายการ Template ทั้งหมด พร้อมข้อมูลเวอร์ชันปัจจุบัน — ใช้แสดงหน้าเว็บ"""
    from customers.cpall.models import TemplateVersion

    result = []
    for key, info in get_template_registry().items():
        _ensure_initial_version(key)
        # Source of truth = Active TemplateGroup
        active = get_active_group_version(key)
        total_versions = TemplateVersion.objects.filter(template_key=key).count()
        result.append({
            "key": key,
            "label": info["label"],
            "exists": active is not None,
            "active_version": active.version_number if active else None,
            "uploaded_at": active.uploaded_at if active else None,
            "total_versions": total_versions,
        })
    return result


def list_versions(key: str) -> list[dict]:
    """คืนทุกเวอร์ชันของ template นี้ เรียงใหม่สุดก่อน — ใช้แสดงหน้าประวัติเวอร์ชัน"""
    from customers.cpall.models import TemplateVersion

    _ensure_initial_version(key)

    versions = (
        TemplateVersion.objects
        .filter(template_key=key)
        .order_by("-version_number")
    )

    result = []

    active_version = get_active_group_version(key)
    active_version_id = active_version.id if active_version else None

    for v in versions:
        reconcile = reconcile_template_with_product_master(
            key=key,
            filepath=v.file_path,
        )

        result.append({
            "id": v.id,
            "version_number": v.version_number,
            "is_active": v.id == active_version_id,
            "uploaded_at": v.uploaded_at,
            "validation_summary": v.validation_summary,
            "original_filename": v.original_filename,

            # ข้อมูล reconciliation
            "reconcile": reconcile,
            "mismatch_count": reconcile["mismatch_count"],
            "is_consistent": reconcile["is_consistent"],
        })

    return result


def upload_new_version(key: str, new_filepath: str, original_filename: str = None) -> dict:
    """
    ตรวจสอบไฟล์ใหม่ก่อน (validate_template) แล้วสร้างเป็นเวอร์ชันใหม่ — แต่ยังไม่ set is_active=True
    ทันที เสมอ (ทั้งกรณีมี mismatch กับ ProductMaster หรือไม่มีก็ตาม) ต้องรอ Admin ยืนยันผ่านหน้า
    reconcile/confirm ก่อนเท่านั้น ถึงจะ activate จริงผ่าน template_version_activate() — ป้องกันไม่ให้
    apply template ใหม่โดยไม่ผ่านการตรวจสอบ ProductMaster เลย เวอร์ชันเก่าไม่ได้ถูกลบเลย แค่ไม่ active
    แล้ว (กู้คืนได้เสมอผ่าน restore_to_version) ถ้า validate ไม่ผ่าน จะ raise ทันที ไม่แตะเวอร์ชัน/ไฟล์
    live เดิมเลย
    original_filename: ชื่อไฟล์ตอน Admin เลือกอัปโหลดจริง (เก็บไว้ให้ดูย้อนหลังในหน้าประวัติเวอร์ชัน)

    *** ใช้กับหน้า Template Versions แบบเดิม (อัปโหลด + reconcile ProductMaster ทันที) เท่านั้น ***
    (2025-09-10) สำหรับ Feature 1 (Template Group) ใช้ upload_version_for_group() แทน — ฟังก์ชันนั้น
    ตรวจแค่โครงสร้างไฟล์ ไม่ reconcile กับ ProductMaster ทันที (defer ไปเช็คตอน Activate Group แทน
    ตาม concept "import เข้า Group ยังไม่เช็คกับ ProductMaster") — เก็บฟังก์ชันนี้ไว้ไม่ลบทิ้ง เผื่อ
    ต้องกลับไปใช้ flow อัปโหลดนอก Group แบบเดิมในอนาคต
    """
    if key not in get_template_registry():
        raise TemplateValidationError(f"ไม่รู้จัก template '{key}'")

    validation_result = validate_template(key, new_filepath)

    from customers.cpall.models import TemplateVersion

    _ensure_initial_version(key)
    last = TemplateVersion.objects.filter(template_key=key).order_by("-version_number").first()
    next_version = (last.version_number + 1) if last else 1

    version_path = _version_file_path(key, next_version)
    os.makedirs(os.path.dirname(version_path), exist_ok=True)
    shutil.move(new_filepath, version_path)

    # ตรวจสอบความสอดคล้องกับ ProductMaster
    reconcile_result = reconcile_template_with_product_master(
        key=key,
        filepath=version_path,
    )

    # ถ้ามี mismatch → สร้าง version แต่ยังไม่ active
    if reconcile_result["mismatch_count"] > 0:
        new_version = TemplateVersion.objects.create(
            customer_id=get_cpall_customer_id(),
            template_key=key,
            version_number=next_version,
            file_path=version_path,
            original_filename=original_filename,
            is_active=False,
            validation_summary=f"sku_count={validation_result.get('sku_count')}",
        )

        return {
            **validation_result,
            "status": "pending",
            "version": new_version,
            "reconcile": reconcile_result,
        }

    # ไม่มี mismatch → ยังไม่ active และยังไม่ sync live
    # รอ Admin กดยืนยันผ่าน popup ก่อน
    new_version = TemplateVersion.objects.create(
        customer_id=get_cpall_customer_id(),
        template_key=key,
        version_number=next_version,
        file_path=version_path,
        original_filename=original_filename,
        is_active=False,
        validation_summary=f"sku_count={validation_result.get('sku_count')}",
    )

    return {
        **validation_result,
        "status": "confirm",
        "version": new_version,
        "reconcile": reconcile_result,
    }


def upload_version_for_group(key: str, new_filepath: str, original_filename: str = None):
    """
    สร้าง TemplateVersion ใหม่สำหรับใช้ใน Template Group (Feature 1, 2025-09-10) — ตรวจแค่โครงสร้าง
    ไฟล์เอง (validate_template) เหมือนกัน แต่ *** ไม่เรียก reconcile_template_with_product_master()
    เลย *** ต่างจาก upload_new_version() เดิมด้านบน — ตาม concept ที่ตกลงกันไว้: "ตอน import ไฟล์เข้า
    Group ยังไม่เช็คกับ ProductMaster" เช็คแค่ตอนกด Activate Group เท่านั้น (ผ่าน
    validate_group_consistency() ก่อน แล้วค่อย reconcile_template_with_product_master() ทีหลัง)

    เวอร์ชันที่สร้างจากฟังก์ชันนี้ยังไม่ active เสมอ (เหมือนกับ upload_new_version()) ต้องถูกเพิ่มเข้า
    Group แล้ว Activate Group ทั้งชุดถึงจะ active จริง — ไม่ sync live file ที่นี่เลย

    คืนค่า TemplateVersion object ที่สร้างใหม่ตรงๆ (ไม่ใช่ dict แบบ upload_new_version() เพราะไม่มี
    reconcile result ให้ส่งกลับ — caller เพิ่มเป็นสมาชิก Group เองด้วย TemplateGroupItem)
    """
    if key not in get_template_registry():
        raise TemplateValidationError(f"ไม่รู้จัก template '{key}'")

    validation_result = validate_template(key, new_filepath)

    from customers.cpall.models import TemplateVersion

    _ensure_initial_version(key)
    last = TemplateVersion.objects.filter(template_key=key).order_by("-version_number").first()
    next_version = (last.version_number + 1) if last else 1

    version_path = _version_file_path(key, next_version)
    os.makedirs(os.path.dirname(version_path), exist_ok=True)
    shutil.move(new_filepath, version_path)

    return TemplateVersion.objects.create(
        customer_id=get_cpall_customer_id(),
        template_key=key,
        version_number=next_version,
        file_path=version_path,
        original_filename=original_filename,
        is_active=False,
        validation_summary=f"sku_count={validation_result.get('sku_count')}",
    )


def restore_to_version(key: str, version_id: int):
    """
    ตรวจสอบ Template Version ที่ต้องการ restore กับ ProductMaster
    ก่อนสลับ Active

    - ไม่มี mismatch → restore และ sync live
    - มี mismatch → ยังไม่ Active และยังไม่ sync live
    """
    from customers.cpall.models import TemplateVersion

    try:
        version = TemplateVersion.objects.get(
            id=version_id,
            template_key=key,
        )
    except TemplateVersion.DoesNotExist:
        raise TemplateValidationError("ไม่พบเวอร์ชันนี้")

    # ตรวจสอบกับ ProductMaster ณ เวลาที่กด restore
    reconcile_result = reconcile_template_with_product_master(
        key=key,
        filepath=version.file_path,
    )

    # มี mismatch → ห้ามเปลี่ยน Active
    if reconcile_result["mismatch_count"] > 0:
        return {
            "status": "pending",
            "version": version,
            "reconcile": reconcile_result,
        }

    # # ไม่มี mismatch → restore ได้
    # TemplateVersion.objects.filter(
    #     template_key=key,
    #     is_active=True,
    # ).update(is_active=False)

    # version.is_active = True
    # version.save(update_fields=["is_active"])

    # _sync_live_file(key, version)

    return {
        "status": "active",
        "version": version,
        "reconcile": reconcile_result,
    }


def delete_version(key: str, version_id: int):
    """
    ลบเวอร์ชันที่ระบุถาวร (ลบทั้ง DB record และไฟล์บนดิสก์) — ลบไม่ได้ถ้า:
      1. เป็นเวอร์ชันสุดท้ายที่เหลืออยู่ของ template นี้ (ต้องมีอย่างน้อย 1 เวอร์ชันเสมอ)
      2. มีแผนที่เคยสร้างไว้ใช้เวอร์ชันนี้อยู่ (เช็คผ่าน plan_run.production_template_version /
         plan_run_logistic_file.template_version — คอลัมน์นี้เพิ่งเตรียมไว้ ยังไม่มีแผนไหนผูกจริง
         จนกว่าจะถึง sub-phase 3 ที่เปลี่ยน flow สร้างแผน ตอนนี้เช็คแล้วจะผ่านเสมอ แต่โค้ดพร้อมใช้)
    ถ้าลบเวอร์ชันที่ active อยู่ จะเลื่อนเวอร์ชันล่าสุดที่เหลือขึ้นมาเป็น active แทนอัตโนมัติ
    """
    from customers.cpall.models import PlanRun, PlanRunLogisticFile, TemplateVersion

    try:
        version = TemplateVersion.objects.get(id=version_id, template_key=key)
    except TemplateVersion.DoesNotExist:
        raise TemplateValidationError("ไม่พบเวอร์ชันนี้")

    total_count = TemplateVersion.objects.filter(template_key=key).count()
    if total_count <= 1:
        raise TemplateInUseError("ลบไม่ได้ — เหลือเทมเพลตนี้แค่เวอร์ชันเดียว ต้องมีอย่างน้อย 1 เวอร์ชันเสมอ")

    used_in_production = PlanRun.objects.filter(production_template_version=version).exists()
    used_in_logistic = PlanRunLogisticFile.objects.filter(template_version=version).exists()
    if used_in_production or used_in_logistic:
        raise TemplateInUseError("ลบไม่ได้ — มีแผนที่เคยสร้างไว้ใช้เทมเพลตเวอร์ชันนี้อยู่")

    active_version = get_active_group_version(key)

    if active_version and active_version.id == version.id:
        raise TemplateInUseError(
            "ลบไม่ได้ — เวอร์ชันนี้กำลังถูกใช้งานอยู่ในชุด Template ที่ Active"
        )

    file_path = version.file_path
    version.delete()

    if os.path.exists(file_path):
        os.remove(file_path)

def get_group_template_versions(group):
    """คืนเฉพาะ TemplateVersion ที่อยู่ใน TemplateGroup นี้"""
    items = (
        group.items
        .select_related("template_version")
        .all()
    )

    production_version = None
    logistic_versions = []

    for item in items:
        version = item.template_version

        if version.template_key == "production_plan":
            production_version = version
        else:
            logistic_versions.append(version)

    return production_version, logistic_versions

def validate_group_consistency(production_version, logistic_versions: list) -> dict:
    """
    ตรวจสอบว่า Production Template กับ Logistic Template ทั้งหมดในชุด (Group) ตรงกันไหม — ตรวจ "ไฟล์
    กับไฟล์" ล้วนๆ ไม่เกี่ยวกับ ProductMaster เลย (Feature 1, 2025-09-10 — เช็คกับ ProductMaster แยก
    ต่างหากตอน Activate Group ผ่าน reconcile_template_with_product_master())

    กติกาที่ตกลงกันไว้:
      1. Product set ของ Production ต้อง = union ของ Product set จาก Logistic ทั้งหมดในชุด
      2. sub_location ของ Production ต้อง = union ของ sub_location จาก Logistic ทั้งหมดในชุด
      3. Logistic แต่ละไฟล์ไม่จำเป็นต้องมีครบทุก sub_location ของ Production (แทนคนละ Logistic Group)
      4. sub_location ซ้ำกันระหว่าง Logistic Templates ต่างไฟล์กัน ยังไม่ห้าม (อาจมี business case)

    production_version, logistic_versions: TemplateVersion objects (ไม่ใช่ key string) — เปิดไฟล์จาก
    version.file_path ตรงๆ (ไม่ใช่ไฟล์ live ที่ path ตายตัว) ตรวจสอบได้แม้ version นั้นยังไม่ active
    """
    import openpyxl

    from customers.cpall.logic.excel_export import _find_sku_header_rows as _find_pp_sku_header_rows
    from customers.cpall.logic.excel_export import _find_sub_location_columns
    from customers.cpall.logic.logistic_plan_export import _find_column_labels, _find_line_no_column
    from customers.cpall.logic.logistic_plan_export import _find_qty_column_range
    from customers.cpall.logic.logistic_plan_export import _find_sku_header_rows as _find_lp_sku_header_rows

    registry = get_template_registry()

    # ---------- Production: barcode set + sub_location set ----------
    if production_version.template_key not in registry or registry[production_version.template_key]["kind"] != "production":
        raise TemplateValidationError(f"'{production_version.template_key}' ไม่ใช่ Production Template")

    pp_wb = openpyxl.load_workbook(production_version.file_path)
    pp_ws = pp_wb[get_pp_sheet_name()]
    pp_headers = _find_pp_sku_header_rows(pp_ws)
    pp_barcodes = set(pp_headers.keys())

    pp_product_names = {
        barcode: str(pp_ws.cell(row=row, column=4).value or "").strip()
        for barcode, row in pp_headers.items()
    }
    pp_sub_locations = set(_find_sub_location_columns(pp_ws).values())

    # ---------- Logistic: union ของ barcode set + sub_location set จากทุกไฟล์ในชุด ----------
    lp_product_names = {}
    lp_barcodes_union = set()
    lp_sub_locations_union = set()

    for lv in logistic_versions:
        if lv.template_key not in registry or registry[lv.template_key]["kind"] != "logistic":
            raise TemplateValidationError(f"'{lv.template_key}' ไม่ใช่ Logistic Template")

        group_name = registry[lv.template_key]["group"]
        _, sheet_name = get_group_templates()[group_name]

        wb = openpyxl.load_workbook(lv.file_path)
        ws = wb[sheet_name]

        line_no_col, header_row = _find_line_no_column(ws)
        name_col = line_no_col + 1
        qty_start_col = line_no_col + 3
        qty_start_col, qty_end_col = _find_qty_column_range(ws, qty_start_col, header_row)

        # forward-fill sub_location ของแต่ละคอลัมน์ — pattern เดียวกับ export_logistic_plan()/
        # validate_logistic_plan() (คอลัมน์ PO2, PO3, ... ที่ตามหลังคอลัมน์แรกของจุดส่งย่อยเดียวกัน
        # ไม่มีชื่อซ้ำเพราะเทมเพลตใช้ merged cell ใส่ชื่อแค่ครั้งเดียว)
        last_sub_location = None
        for col in range(qty_start_col, qty_end_col + 1):
            sub_loc, _ = _find_column_labels(ws, col, header_row)
            if sub_loc is None:
                sub_loc = last_sub_location if last_sub_location is not None else group_name
            last_sub_location = sub_loc
            lp_sub_locations_union.add(sub_loc)

        header_rows = _find_lp_sku_header_rows(ws, name_col)
        lp_barcodes_union |= set(header_rows.keys())

        for barcode, row in header_rows.items():
            lp_product_names.setdefault(
                barcode,
                str(ws.cell(row=row, column=name_col).value or "").strip(),
            )
    # ---------- เปรียบเทียบ ----------
    product_missing_in_logistic = [
        {
            "barcode": barcode,
            "item_name": pp_product_names.get(barcode, ""),
        }
        for barcode in sorted(pp_barcodes - lp_barcodes_union)
    ]

    product_extra_in_logistic = [
        {
            "barcode": barcode,
            "item_name": lp_product_names.get(barcode, ""),
        }
        for barcode in sorted(lp_barcodes_union - pp_barcodes)
    ]
    sub_location_missing_in_logistic = sorted(pp_sub_locations - lp_sub_locations_union)
    sub_location_extra_in_logistic = sorted(lp_sub_locations_union - pp_sub_locations)

    is_consistent = not (
        product_missing_in_logistic or product_extra_in_logistic
        or sub_location_missing_in_logistic or sub_location_extra_in_logistic
    )

    return {
        "is_consistent": is_consistent,
        "product_missing_in_logistic": product_missing_in_logistic,
        "product_extra_in_logistic": product_extra_in_logistic,
        "sub_location_missing_in_logistic": sub_location_missing_in_logistic,
        "sub_location_extra_in_logistic": sub_location_extra_in_logistic,
    }


def ensure_bootstrap_group():
    """
    Bootstrap "ชุดปัจจุบัน" (Feature 1, Step 7, 2025-09-11) — ถ้ายังไม่มี TemplateGroup เลยในระบบ
    (ระบบเพิ่งอัปเกรดมาใช้ Template Group ครั้งแรก) สร้าง Group แรกอัตโนมัติจากเวอร์ชันที่ active
    อยู่ตอนนี้ทั้งหมด — ไม่ต้องให้ Admin ทำอะไรเพิ่ม ป้องกันข้อมูลสับสน (ถ้าไม่ทำ Admin จะอัปโหลด/
    เปลี่ยนเวอร์ชันอะไรไม่ได้เลยจนกว่าจะสร้าง Group เอง เพราะการอัปโหลด/activate ทั้งหมดย้ายมาอยู่ใต้
    Template Group แล้ว — Step 6) เรียกแบบ idempotent เหมือนกับ _ensure_initial_version() ด้านบน —
    เรียกซ้ำกี่ครั้งก็ปลอดภัย เพราะเช็ค TemplateGroup.objects.exists() ก่อนเสมอ
    """
    from customers.cpall.models import TemplateGroup, TemplateGroupItem, TemplateVersion

    if TemplateGroup.objects.exists():
        return

    active_versions = list(TemplateVersion.objects.filter(is_active=True))
    if not active_versions:
        return  # fresh install ที่ยังไม่เคยอัปโหลด template อะไรเลย ไม่มีอะไรให้ bootstrap

    from django.utils import timezone

    group = TemplateGroup.objects.create(
        customer_id=get_cpall_customer_id(), name="ชุดปัจจุบัน",
        note="สร้างอัตโนมัติตอนอัปเกรดมาใช้ Template Group ครั้งแรก จากเวอร์ชันที่ใช้งานอยู่ก่อนหน้านี้",
        is_active=True, activated_at=timezone.now(),
    )
    for v in active_versions:
        TemplateGroupItem.objects.create(template_group=group, template_version=v)



