"""
plan_runner.py — ตัวกลางที่รวม logic การสร้างแผนไว้ที่เดียว

เดิม logic นี้อยู่ใน main.py cmd_plan() ปนกับ print() ทำให้เอาไปใช้กับหน้าเว็บไม่ได้ (เว็บต้องการ
ข้อมูลเป็น dict ไปแสดงผล ไม่ใช่ข้อความ terminal) — ย้ายมาไว้ที่นี่ คืนค่าเป็น dict ล้วนๆ
ส่วน main.py และเว็บ ต่างคนต่าง format ผลลัพธ์เป็นข้อความ/HTML ของตัวเอง

บันทึกประวัติการสร้างแผนลงตาราง plan_run ทุกครั้งที่เรียก (ไม่ว่าจะสำเร็จหรือมีบางไฟล์ล้มเหลว)
เพื่อให้หน้าเว็บย้อนกลับมาดูได้ว่าเคยสร้างแผนอะไรไว้บ้าง

*** Phase 1 (ORM migration) *** ฟังก์ชัน CRUD ในไฟล์นี้ (บันทึก/อ่าน/ลบ plan_run) ย้ายมาใช้ Django ORM
แล้ว (เดิมเป็น raw SQL ทั้งหมด) — run_plan() เองยังเรียก reconcile()/export_*() ที่เป็น raw SQL อยู่
(query คำนวณซับซ้อนพวกนั้นตั้งใจไม่ย้าย ดู grouping.py) function signature/return shape เดิมทุกตัว
ไม่เปลี่ยน เพื่อไม่ให้ต้องแก้ views.py เลย
"""
import os
import shutil
from datetime import datetime

from customers.cpall.logic.db import get_cpall_customer_id
from customers.cpall.logic.excel_export import ExcelExportError, export_production_plan
from customers.cpall.logic.grouping import ReconciliationError, reconcile
from customers.cpall.logic.logistic_plan_export import (
    LogisticPlanError,
    export_logistic_plan,
    get_group_templates,
    group_has_data,
)
from customers.cpall.models import PlanRun, PlanRunLogisticFile


def run_plan(po_import_ids: list[int], output_dir: str | None = None, buffer_override: dict = None) -> dict:
    """
    รัน pipeline สร้างแผนทั้งหมด (reconcile -> Production Plan -> Logistic Plan 4 กลุ่ม)
    คืนค่าเป็น dict ล้วนๆ ไม่ print — ผู้เรียก (main.py / เว็บ) เอาไป format แสดงผลเอง

    ถ้ายอดไม่ตรง (reconciliation ล้มเหลว) จะ raise ReconciliationError ทันที ไม่สร้างไฟล์ใดๆ เลย
    (ไม่บันทึก plan_run ด้วย เพราะยังไม่มีอะไรให้บันทึก)

    ถ้า reconcile ผ่านแต่บางไฟล์ (Production Plan หรือ Logistic Plan บางกลุ่ม) ล้มเหลวระหว่างสร้าง
    จะไม่ raise — ใส่สถานะ "failed" ไว้ใน dict ผลลัพธ์แทน (ให้ไฟล์อื่นที่ไม่เกี่ยวข้องสร้างต่อได้)
    """
    # ---------- ตรวจสอบยอด ----------
    # เช็คก่อนสร้างโฟลเดอร์เลย — ถ้ายอดไม่ตรง (raise ทันที) จะได้ไม่มีโฟลเดอร์เปล่าค้างอยู่บนดิสก์
    recon = reconcile(po_import_ids)  # raise ReconciliationError ถ้าไม่ผ่าน (ปล่อยให้ผู้เรียนจัดการ)
    if not recon["passed"]:
        raise ReconciliationError(
            f"ยอดไม่ตรงกัน {len(recon['mismatches'])} SKU: "
            + ", ".join(m["barcode"] for m in recon["mismatches"])
        )

    if output_dir is None:
        timestamp = datetime.now().strftime("%Y%m%d_%H%M")
        output_dir = f"customers/cpall/data/output/{timestamp}"
    os.makedirs(output_dir, exist_ok=True)

    # ---------- Production Plan ----------
    production_plan_path = f"{output_dir}/แพลน_7-11.xlsx"
    production_plan_result = {"status": "success", "path": production_plan_path, "error": None}
    try:
        export_production_plan(po_import_ids, production_plan_path, buffer_override=buffer_override)
    except (ExcelExportError, Exception) as e:
        production_plan_result = {"status": "failed", "path": None, "error": str(e)}

    # ---------- Logistic Plan (แต่ละกลุ่ม อิสระต่อกัน) ----------
    logistic_results = {}
    for group_name in get_group_templates():
        if not group_has_data(po_import_ids, group_name):
            logistic_results[group_name] = {"status": "skipped", "path": None, "error": None}
            continue

        group_output_path = f"{output_dir}/{group_name}.xlsx"
        try:
            export_logistic_plan(po_import_ids, group_name, group_output_path)
            logistic_results[group_name] = {"status": "success", "path": group_output_path, "error": None}
        except (LogisticPlanError, Exception) as e:
            logistic_results[group_name] = {"status": "failed", "path": None, "error": str(e)}

    # ---------- บันทึกประวัติลง DB (ผ่าน ORM) ----------
    plan_run_id = _save_plan_run(po_import_ids, output_dir, production_plan_result, logistic_results)

    # ---------- ดึงผลลัพธ์จริง (LibreOffice คำนวณสูตรจริง) เก็บลง plan_sku_result ----------
    # (Phase 1.6 sub-phase 3) — ทำแยกหลัง save plan_run เสร็จแล้ว ถ้าขั้นตอนนี้ error ไม่ควรทำให้
    # การสร้างแผนทั้งหมดล้มเหลว (ไฟล์ Excel สร้างสำเร็จแล้ว ยังใช้งานได้ปกติ แค่ตาราง data-first
    # จะไม่มีข้อมูลสำหรับแผนนี้ — log ไว้ให้เห็นชัดเจนแทนที่จะทำให้ทั้ง request ล้ม)
    extracted_ok = _extract_and_save_sku_results(plan_run_id, production_plan_result, logistic_results)

    # ---------- ลบไฟล์ Excel ที่ extract ข้อมูลเข้า plan_sku_result สำเร็จแล้วทิ้ง (data-first เต็มรูป
    # แบบ) — เว็บอ่านจาก DB อยู่แล้ว ดาวน์โหลดก็ regenerate จาก DB+เทมเพลตใหม่ทุกครั้งอยู่แล้ว ไฟล์ที่
    # เขียนไว้ตอนสร้างแผนไม่มีประโยชน์อะไรอีก — ไฟล์ที่ extract ไม่สำเร็จ (หายาก) เก็บไว้เป็น fallback
    # ให้ดาวน์โหลดยังใช้งานได้ ไม่ลบทิ้ง
    if "production" in extracted_ok and production_plan_result["path"]:
        if os.path.exists(production_plan_result["path"]):
            os.remove(production_plan_result["path"])
    for group_name, result in logistic_results.items():
        if group_name in extracted_ok and result["path"] and os.path.exists(result["path"]):
            os.remove(result["path"])

    # โฟลเดอร์นี้อาจว่างเปล่าแล้ว (ทุกไฟล์ที่เคยเขียนไว้ถูกลบไปหมด) — ลบทิ้งไปด้วยเลยถ้าว่างจริง กัน
    # โฟลเดอร์เปล่าค้างสะสมไปเรื่อยๆ (ถ้ายังมีไฟล์เหลืออยู่ — extract ไม่สำเร็จบางไฟล์ — จะไม่ว่าง เลย
    # ไม่ถูกลบ ปลอดภัย)
    if os.path.isdir(output_dir) and not os.listdir(output_dir):
        os.rmdir(output_dir)

    return {
        "plan_run_id": plan_run_id,
        "po_import_ids": po_import_ids,
        "output_dir": output_dir,
        "reconciliation": recon,
        "production_plan": production_plan_result,
        "logistic_plans": logistic_results,
    }


def _extract_and_save_sku_results(plan_run_id, production_plan_result, logistic_results) -> set:
    """ดึงผลลัพธ์ (LibreOffice คำนวณจริง) เก็บลง plan_sku_result — คืนค่า set ของไฟล์ที่ extract
    สำเร็จ ('production' หรือชื่อกลุ่ม) ให้ run_plan() ใช้ตัดสินใจว่าจะลบไฟล์ Excel ทิ้งได้ไหม (ไฟล์ที่
    extract ไม่สำเร็จต้องเก็บไว้เป็น fallback ให้ดาวน์โหลด — ไม่งั้นข้อมูลจะหายไปเลยทั้งสองทาง)"""
    from customers.cpall.logic.plan_result_extractor import (
        extract_logistic_plan_results,
        extract_production_plan_results,
    )
    from customers.cpall.models import PlanSkuResult

    all_rows = []
    extracted_ok = set()

    if production_plan_result["status"] == "success" and production_plan_result["path"]:
        try:
            all_rows += extract_production_plan_results(production_plan_result["path"])
            extracted_ok.add("production")
        except Exception as e:
            print(f"[plan_runner] WARNING: ดึงผลลัพธ์ Production Plan ไม่สำเร็จ (plan_run_id={plan_run_id}): {e}")

    for group_name, result in logistic_results.items():
        if result["status"] == "success" and result["path"]:
            try:
                all_rows += extract_logistic_plan_results(result["path"], group_name)
                extracted_ok.add(group_name)
            except Exception as e:
                print(f"[plan_runner] WARNING: ดึงผลลัพธ์ {group_name} ไม่สำเร็จ (plan_run_id={plan_run_id}): {e}")

    if all_rows:
        PlanSkuResult.objects.bulk_create([
            PlanSkuResult(plan_run_id=plan_run_id, **row) for row in all_rows
        ])
        print(f"[plan_runner] บันทึกผลลัพธ์ลง plan_sku_result แล้ว {len(all_rows)} แถว (plan_run_id={plan_run_id})")

    return extracted_ok


def _save_plan_run(po_import_ids, output_dir, production_plan_result, logistic_results) -> int:
    from customers.cpall.logic.template_manager import _ensure_initial_version
    from customers.cpall.models import TemplateVersion

    customer_id = get_cpall_customer_id()

    # ผูกแผนเข้ากับเทมเพลตเวอร์ชันที่ active ตอนสร้างแผน (Phase 1.6 sub-phase 5) — ให้รู้ย้อนหลังได้ว่า
    # แผนนี้ใช้เทมเพลตเวอร์ชันไหน (ดาวน์โหลดซ้ำจะได้ตัวเลขตรงกับตอนสร้างแม้เทมเพลตปัจจุบันถูกแก้ไปแล้ว)
    # และให้กฎ "ลบเทมเพลตเวอร์ชันที่มีแผนอ้างอิงอยู่ไม่ได้" มีผลจริง — เรียก _ensure_initial_version
    # ก่อนเสมอ กันเคส "ยังไม่เคยมีใครเปิดหน้า Template เลย" ซึ่งจะยังไม่มี TemplateVersion แถวไหนเลย
    # หา TemplateVersion ไม่เจอจริงๆ (เช่นไม่มีไฟล์ live เลย) ก็ไม่เป็นไร (NULL ได้ ไม่บล็อกงานหลัก)
    _ensure_initial_version("production_plan")
    production_template_version = TemplateVersion.objects.filter(
        template_key="production_plan", is_active=True
    ).first()

    plan_run = PlanRun.objects.create(
        customer_id=customer_id,
        output_dir=output_dir,
        production_plan_path=production_plan_result["path"],
        production_plan_status=production_plan_result["status"],
        production_plan_error=production_plan_result["error"],
        production_template_version=production_template_version,
    )
    plan_run.po_imports.set(po_import_ids)  # ORM จัดการ INSERT เข้า plan_run_import (through table) ให้เอง

    logistic_file_objs = []
    for group_name, result in logistic_results.items():
        key = f"logistic_{group_name}"
        _ensure_initial_version(key)
        template_version = TemplateVersion.objects.filter(template_key=key, is_active=True).first()
        logistic_file_objs.append(PlanRunLogisticFile(
            plan_run=plan_run, group_name=group_name, status=result["status"],
            file_path=result["path"], error_message=result["error"],
            template_version=template_version,
        ))
    PlanRunLogisticFile.objects.bulk_create(logistic_file_objs)

    return plan_run.id


def list_plan_runs_paginated(page: int = 1, page_size: int = 10, status: str = "all",
                              date_from=None, date_to=None) -> dict:
    """
    ดึงรายการแผนแบบแบ่งหน้า + filter ตามสถานะ/ช่วงวันที่สร้างได้ — ใช้ที่หน้า "แผนทั้งหมด"
    status: "all" / "success" / "failed"
    date_from, date_to: datetime.date หรือ None — กรองจาก created_at (วันที่สร้างแผน)
    -> {"items": [...], "total": N, "page": ..., "page_size": ..., "total_pages": ...}
    """
    page = max(1, page)
    page_size = page_size if page_size in (10, 50, 100) else 10
    offset = (page - 1) * page_size

    qs = PlanRun.objects.all()
    if status == "success":
        qs = qs.filter(production_plan_status="success")
    elif status == "failed":
        qs = qs.exclude(production_plan_status="success")
    if date_from:
        qs = qs.filter(created_at__date__gte=date_from)
    if date_to:
        qs = qs.filter(created_at__date__lte=date_to)

    total = qs.count()
    rows = qs.order_by("-created_at")[offset:offset + page_size]

    items = [
        {"id": r.id, "created_at": r.created_at, "output_dir": r.output_dir,
         "production_plan_status": r.production_plan_status}
        for r in rows
    ]
    total_pages = max(1, (total + page_size - 1) // page_size)
    return {"items": items, "total": total, "page": page, "page_size": page_size, "total_pages": total_pages}


def list_plan_runs(limit: int = 50) -> list[dict]:
    """ดึงรายการแผนที่เคยสร้างไว้ทั้งหมด เรียงล่าสุดก่อน — ใช้แสดงในหน้า Dashboard"""
    rows = PlanRun.objects.order_by("-created_at")[:limit]
    return [
        {"id": r.id, "created_at": r.created_at, "output_dir": r.output_dir,
         "production_plan_status": r.production_plan_status}
        for r in rows
    ]


def get_plan_run_detail(plan_run_id: int) -> dict | None:
    """ดึงรายละเอียดแผนที่สร้างไว้ 1 รายการ (ใช้แสดงหน้า ดูแผน)"""
    try:
        plan_run = PlanRun.objects.get(id=plan_run_id)
    except PlanRun.DoesNotExist:
        return None

    result = {
        "id": plan_run.id, "created_at": plan_run.created_at, "output_dir": plan_run.output_dir,
        "production_plan_path": plan_run.production_plan_path,
        "production_plan_status": plan_run.production_plan_status,
        "production_plan_error": plan_run.production_plan_error,
    }
    result["po_imports"] = [
        {"id": pi.id, "source_filename": pi.source_filename,
         "display_filename": os.path.basename(pi.source_filename), "production_date": pi.production_date,
         "po_date": pi.po_date}
        for pi in plan_run.po_imports.all()
    ]
    result["logistic_plans"] = [
        {"group_name": lf.group_name, "status": lf.status, "file_path": lf.file_path,
         "error_message": lf.error_message}
        for lf in plan_run.logistic_files.order_by("group_name")
    ]
    return result


def delete_plan_run(plan_run_id: int):
    """
    ลบแผนที่สร้างไว้ (ลบทั้งประวัติใน DB และไฟล์ Excel ทั้งหมดของแผนนั้นบนดิสก์)
    ไม่กระทบ po_import ต้นทางเลย (แผนแค่ "อ้างอิง" po_import ไม่ได้เป็นเจ้าของข้อมูล) — ลบแผนได้อิสระ
    plan_run_import / plan_run_logistic_file ตั้ง on_delete=CASCADE ไว้ใน model แล้ว ลบตามอัตโนมัติ
    """
    try:
        plan_run = PlanRun.objects.get(id=plan_run_id)
    except PlanRun.DoesNotExist:
        return

    output_dir = plan_run.output_dir
    plan_run.delete()

    if output_dir and os.path.isdir(output_dir):
        shutil.rmtree(output_dir)
