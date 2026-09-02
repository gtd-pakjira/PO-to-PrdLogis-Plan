"""
views.py — หน้าเว็บทั้งหมดของโมดูล cpall (CP All / 7-11)
"""
import json
import os
from datetime import datetime

from django.http import FileResponse, Http404, HttpResponse
from django.shortcuts import redirect, render
from django.urls import reverse

from customers.cpall.forms import ImportPOForm, TemplateUploadForm
from customers.cpall.logic.grouping import ReconciliationError
from customers.cpall.logic.location_mapping_manager import get_existing_groups, save_location_mapping
from customers.cpall.logic.plan_regenerator import (
    PlanRegenerateError,
    regenerate_logistic_plan_bytes,
    regenerate_production_plan_bytes,
)
from customers.cpall.logic.plan_runner import (
    delete_plan_run,
    get_plan_run_detail,
    list_plan_runs,
    list_plan_runs_paginated,
    run_plan,
)
from customers.cpall.logic.plan_view_data import (
    get_logistic_plan_table,
    get_logistic_plan_table_from_db,
    get_production_plan_table,
    get_production_plan_table_from_db,
)
from customers.cpall.logic.po_parser import (
    POInUseError,
    POParseError,
    check_unknown_locations,
    delete_po_import,
    list_po_imports,
    list_po_imports_paginated,
    load_po_to_db,
)
from customers.cpall.logic.template_manager import (
    TEMPLATE_REGISTRY,
    TemplateInUseError,
    TemplateValidationError,
    delete_version,
    get_template_grid,
    list_templates,
    list_versions,
    restore_to_version,
    upload_new_version,
)
from customers.cpall.models import PlanRun

UPLOAD_DIR = "customers/cpall/data/po_uploads"
TEMP_UPLOAD_DIR = "customers/cpall/data/temp_uploads"


def index(request):
    """แดชบอร์ด — สรุปภาพรวมสั้นๆ เท่านั้น (3 PO ล่าสุด + 3 แผนล่าสุด) ดูทั้งหมดแยกไปหน้า /po/ และ /plans/"""
    po_imports = list_po_imports(limit=3)
    plan_runs = list_plan_runs(limit=3)
    return render(request, "cpall/index.html", {"po_imports": po_imports, "plan_runs": plan_runs})


def po_list(request):
    """หน้ารายการ PO ทั้งหมด — นำเข้าใหม่/ติ๊กเลือกสร้างแผน/ลบ/ค้นหา/แบ่งหน้า ได้จากหน้านี้"""
    from urllib.parse import urlencode
    page = int(request.GET.get("page", 1) or 1)
    page_size = int(request.GET.get("page_size", 10) or 10)
    search = request.GET.get("q", "").strip()
    result = list_po_imports_paginated(page=page, page_size=page_size, search=search)
    base_qs = ("&" + urlencode({"q": search})) if search else ""
    return render(request, "cpall/po_list.html", {
        "po_imports": result["items"], "total": result["total"], "page": result["page"],
        "page_size": result["page_size"], "total_pages": result["total_pages"],
        "search": search, "base_qs": base_qs,
    })


def plan_list(request):
    """หน้ารายการแผนทั้งหมด — กด "สร้างแผนใหม่" แล้วไปเลือก PO ที่หน้า /po/"""
    from customers.cpall.logic.date_utils import parse_date_arg

    page = int(request.GET.get("page", 1) or 1)
    page_size = int(request.GET.get("page_size", 10) or 10)
    status = request.GET.get("status", "all")
    date_from_str = request.GET.get("date_from", "").strip()
    date_to_str = request.GET.get("date_to", "").strip()
    try:
        date_from = parse_date_arg(date_from_str) if date_from_str else None
    except ValueError:
        date_from = None
    try:
        date_to = parse_date_arg(date_to_str) if date_to_str else None
    except ValueError:
        date_to = None

    result = list_plan_runs_paginated(page=page, page_size=page_size, status=status,
                                       date_from=date_from, date_to=date_to)
    base_qs = f"&status={status}" if status != "all" else ""
    if date_from_str:
        base_qs += f"&date_from={date_from_str}"
    if date_to_str:
        base_qs += f"&date_to={date_to_str}"
    return render(request, "cpall/plan_list.html", {
        "plan_runs": result["items"], "total": result["total"], "page": result["page"],
        "page_size": result["page_size"], "total_pages": result["total_pages"],
        "status": status, "base_qs": base_qs,
        "date_from": date_from_str, "date_to": date_to_str,
    })


def import_form(request):
    return render(request, "cpall/import.html", {"form": ImportPOForm()})


def import_submit(request):
    if request.method != "POST":
        return redirect("cpall:import_form")

    is_htmx = request.headers.get("HX-Request") == "true"

    form = ImportPOForm(request.POST, request.FILES)
    if not form.is_valid():
        # Django Form validation ไม่ผ่าน (กรอกไม่ครบ/ไฟล์ผิดชนิด) -> อยู่หน้าเดิม โชว์ error ในฟอร์มเลย
        # ผ่าน HTMX swap แค่ตัวฟอร์ม (_import_form.html) ไม่ reload ทั้งหน้า
        if is_htmx:
            return render(request, "cpall/_import_form.html", {"form": form})
        return render(request, "cpall/import.html", {"form": form})

    po_file = form.cleaned_data["po_file"]
    production_date = form.cleaned_data["production_date"]
    po_date = form.cleaned_data["po_date"]

    os.makedirs(UPLOAD_DIR, exist_ok=True)
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    original_name = os.path.basename(po_file.name)  # กันเผื่อชื่อไฟล์มี path แปลกปน
    saved_path = os.path.join(UPLOAD_DIR, f"{timestamp}_{original_name}")
    with open(saved_path, "wb") as f:
        for chunk in po_file.chunks():
            f.write(chunk)

    def error_response(message, status=400):
        # นำเข้าไม่สำเร็จ -> ไม่เก็บไฟล์ที่เพิ่งบันทึกไว้ค้างบนดิสก์ (กันพื้นที่บวมจากไฟล์ที่ import พลาด)
        if os.path.exists(saved_path):
            os.remove(saved_path)
        if is_htmx:
            response = HttpResponse(status=status)
            response["HX-Trigger"] = json.dumps({"toast": {"message": message, "level": "error"}})
            return response
        return render(request, "cpall/import_result.html",
                       {"success": False, "error": message, "original_filename": po_file.name})

    try:
        po_import_id = load_po_to_db(saved_path, production_date, po_date, imported_by="web")
    except POParseError as e:
        return error_response(f"ไฟล์ PO มีปัญหา: {e}")
    except Exception as e:
        return error_response(f"นำเข้าล้มเหลว: {type(e).__name__}: {e}", status=500)

    unknown_locations = check_unknown_locations(po_import_id)
    if unknown_locations:
        # ไม่ใช่ dead-end แล้ว — พาไปหน้าเลือก mapping เลย (ข้อมูล PO import สำเร็จแล้วจริงๆ ใน DB
        # แค่ยังมีรหัสสถานที่ที่ไม่รู้จักกลุ่มพื้นที่ ให้ Admin เลือกตรงนี้ได้เลย)
        if is_htmx:
            response = HttpResponse(status=200)
            response["HX-Redirect"] = reverse("cpall:resolve_locations", args=[po_import_id])
            return response
        return redirect("cpall:resolve_locations", po_import_id=po_import_id)

    # สำเร็จสมบูรณ์ -> ไปหน้า PO ทั้งหมดเลย (เห็นผลลัพธ์อยู่ในบริบทของรายการจริง แทนที่จะเป็นหน้า
    # สรุปผลโดดๆ ที่ต้องกดออกไปอีกที)
    if is_htmx:
        response = HttpResponse(status=200)
        response["HX-Redirect"] = reverse("cpall:po_list")
        return response
    return redirect("cpall:po_list")


def resolve_locations(request, po_import_id):
    unknown_locations = check_unknown_locations(po_import_id)
    is_htmx = request.headers.get("HX-Request") == "true"

    if request.method == "POST":
        for fc_code, name_th in unknown_locations:
            group = request.POST.get(f"group_{fc_code}", "").strip()
            new_group = request.POST.get(f"new_group_{fc_code}", "").strip()
            sub_location = request.POST.get(f"sub_location_{fc_code}", "").strip()

            final_group = new_group if new_group else group
            if final_group and sub_location:
                save_location_mapping(fc_code, name_th, final_group, sub_location)

        remaining = check_unknown_locations(po_import_id)
        if remaining:
            # บางรายการยังไม่ได้กรอก (เว้นว่างไว้) -> โชว์ฟอร์มเดิมต่อ เฉพาะรายการที่เหลือ
            context = {
                "po_import_id": po_import_id, "unknown_locations": remaining,
                "existing_groups": get_existing_groups(),
                "error": "ยังมีบางรายการที่ยังไม่ได้เลือกกลุ่มพื้นที่/กรอกจุดส่งย่อย",
            }
            if is_htmx:
                return render(request, "cpall/_resolve_locations_form.html", context)
            return render(request, "cpall/resolve_locations.html", context)

        if is_htmx:
            response = HttpResponse(status=200)
            response["HX-Redirect"] = reverse("cpall:po_list")
            return response
        return redirect("cpall:po_list")

    return render(request, "cpall/resolve_locations.html", {
        "po_import_id": po_import_id, "unknown_locations": unknown_locations,
        "existing_groups": get_existing_groups(),
    })


def new_plan_submit(request):
    if request.method != "POST":
        return redirect("cpall:index")

    is_htmx = request.headers.get("HX-Request") == "true"

    def error_response(message, status=400):
        if is_htmx:
            response = HttpResponse(status=status)
            response["HX-Trigger"] = json.dumps({"toast": {"message": message, "level": "error"}})
            return response
        return render(request, "cpall/plan_error.html", {"error": message})

    po_import_ids = [int(x) for x in request.POST.getlist("po_import_ids")]
    if not po_import_ids:
        return error_response("ต้องเลือก PO อย่างน้อย 1 รอบ")

    # ยอดเผื่อเกี่ยวข้องกับกลุ่ม "รอบเช้าต่างจังหวัด" เท่านั้น (เทมเพลตกลุ่มอื่นไม่มีคอลัมน์นี้เลย) —
    # ถามเฉพาะตอนที่รอบ PO ที่เลือกมีข้อมูลกลุ่มนี้จริง ไม่งั้นข้ามหน้ากรอกยอดเผื่อไปสร้างแผนได้เลย
    # (ยอดเผื่อยังไม่มีสูตรคำนวณที่แน่นอน รอถาม Admin ก่อน — ดู README)
    from customers.cpall.logic.logistic_plan_export import group_has_data
    if group_has_data(po_import_ids, "รอบเช้าต่างจังหวัด"):
        # ไปหน้ากรอกยอดเผื่อจริง (ไม่ใช่ error) — ต้อง navigate เต็มหน้าเสมอ ไม่ใช่แค่เด้ง popup เฉยๆ
        if is_htmx:
            response = HttpResponse(status=200)
            response["HX-Redirect"] = _buffer_form_url(po_import_ids)
            return response
        return redirect_to_buffer_form(request, po_import_ids)

    try:
        result = run_plan(po_import_ids, buffer_override={})  # ไม่มีรอบเช้าต่างจังหวัด -> ไม่ใส่ยอดเผื่อเลย
    except ReconciliationError as e:
        return error_response(str(e), status=409)
    except Exception as e:
        return error_response(f"สร้างแผนล้มเหลว: {type(e).__name__}: {e}", status=500)

    if is_htmx:
        response = HttpResponse(status=200)
        response["HX-Redirect"] = reverse("cpall:view_plan", args=[result["plan_run_id"]])
        return response
    return redirect("cpall:view_plan", plan_run_id=result["plan_run_id"])


def _buffer_form_url(po_import_ids):
    from urllib.parse import urlencode
    qs = urlencode({"po_import_ids": ",".join(str(x) for x in po_import_ids)})
    return f"{reverse('cpall:buffer_form')}?{qs}"


def redirect_to_buffer_form(request, po_import_ids):
    return redirect(_buffer_form_url(po_import_ids))


def buffer_form(request):
    import openpyxl

    from customers.cpall.logic.logistic_plan_export import (
        GROUP_TEMPLATES,
        _find_line_no_column,
        _find_sku_header_rows,
        read_buffer_qty_from_template,
    )
    from customers.cpall.models import SkuMaster

    po_import_ids_str = request.GET.get("po_import_ids", "")
    po_import_ids = [int(x) for x in po_import_ids_str.split(",") if x]
    if not po_import_ids:
        return render(request, "cpall/plan_error.html", {"error": "ไม่พบรอบ PO ที่เลือกไว้"})

    # กรอกยอดเผื่อเฉพาะ SKU ที่มีอยู่ในเทมเพลต "รอบเช้าต่างจังหวัด" เท่านั้น (18 ตัว ไม่ใช่ทั้ง 19)
    # เพราะยอดเผื่อผูกกับกลุ่มนี้กลุ่มเดียวในทางปฏิบัติ (ดูเหตุผลใน README)
    template_path, sheet_name = GROUP_TEMPLATES["รอบเช้าต่างจังหวัด"]
    wb = openpyxl.load_workbook(template_path)
    ws = wb[sheet_name]
    line_no_col, header_row = _find_line_no_column(ws)
    name_col = line_no_col + 1
    barcodes = list(_find_sku_header_rows(ws, name_col).keys())

    try:
        default_buffer = read_buffer_qty_from_template()
    except Exception:
        default_buffer = {}

    name_lookup = {s.barcode: s.name_th for s in SkuMaster.objects.filter(barcode__in=barcodes)}
    sku_rows = sorted(
        [
            {"barcode": bc, "name_th": name_lookup.get(bc, bc), "default_buffer": default_buffer.get(bc, 0)}
            for bc in barcodes
        ],
        key=lambda r: r["name_th"],
    )

    return render(request, "cpall/buffer_form.html", {
        "po_import_ids": po_import_ids,
        "po_import_ids_str": po_import_ids_str,
        "sku_rows": sku_rows,
    })


def buffer_form_submit(request):
    if request.method != "POST":
        return redirect("cpall:po_list")

    is_htmx = request.headers.get("HX-Request") == "true"

    def error_response(message, status=400):
        if is_htmx:
            response = HttpResponse(status=status)
            response["HX-Trigger"] = json.dumps({"toast": {"message": message, "level": "error"}})
            return response
        return render(request, "cpall/plan_error.html", {"error": message})

    po_import_ids_str = request.POST.get("po_import_ids_str", "")
    po_import_ids = [int(x) for x in po_import_ids_str.split(",") if x]
    if not po_import_ids:
        return error_response("ไม่พบรอบ PO ที่เลือกไว้")

    buffer_override = {}
    for key, val in request.POST.items():
        if key.startswith("buffer_") and val.strip():
            barcode = key[len("buffer_"):]
            try:
                buffer_override[barcode] = float(val)
            except ValueError:
                pass

    try:
        result = run_plan(po_import_ids, buffer_override=buffer_override)
    except ReconciliationError as e:
        return error_response(str(e), status=409)
    except Exception as e:
        return error_response(f"สร้างแผนล้มเหลว: {type(e).__name__}: {e}", status=500)

    if is_htmx:
        response = HttpResponse(status=200)
        response["HX-Redirect"] = reverse("cpall:view_plan", args=[result["plan_run_id"]])
        return response
    return redirect("cpall:view_plan", plan_run_id=result["plan_run_id"])


def view_plan(request, plan_run_id):
    detail = get_plan_run_detail(plan_run_id)
    if detail is None:
        return render(request, "cpall/plan_not_found.html", {"plan_run_id": plan_run_id}, status=404)
    plan_name = PlanRun.objects.get(id=plan_run_id).get_display_name()
    return render(request, "cpall/plan_view.html", {"plan": detail, "plan_name": plan_name})


def download_production(request, plan_run_id):
    detail = get_plan_run_detail(plan_run_id)
    if detail is None or not detail["production_plan_path"]:
        raise Http404
    plan_name = PlanRun.objects.get(id=plan_run_id).get_display_name()

    # ลองสร้างไฟล์ใหม่จากข้อมูลดิบก่อนเสมอ (มีสูตรจริงครบ ตรงกับที่ตกลงกันไว้) — ถ้าทำไม่ได้ (แผนเก่า
    # ก่อนมีระบบ data-first) fallback ไปเสิร์ฟไฟล์ที่เก็บไว้บนดิสก์แบบเดิม (ยังไม่ได้ลบไฟล์เก่าทิ้งในเฟสนี้)
    try:
        content = regenerate_production_plan_bytes(plan_run_id)
        response = HttpResponse(
            content,
            content_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
        response["Content-Disposition"] = f'attachment; filename="{plan_name}.xlsx"'
        return response
    except PlanRegenerateError:
        pass

    return FileResponse(open(detail["production_plan_path"], "rb"), as_attachment=True,
                         filename=f"{plan_name}.xlsx")


def download_logistic(request, plan_run_id, group_name):
    detail = get_plan_run_detail(plan_run_id)
    if detail is None:
        raise Http404
    match = next((lp for lp in detail["logistic_plans"] if lp["group_name"] == group_name), None)
    if match is None or not match["file_path"]:
        raise Http404
    plan_name = PlanRun.objects.get(id=plan_run_id).get_display_name()
    filename = f"{group_name}_{plan_name}.xlsx"

    try:
        content = regenerate_logistic_plan_bytes(plan_run_id, group_name)
        response = HttpResponse(
            content,
            content_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
        response["Content-Disposition"] = f'attachment; filename="{filename}"'
        return response
    except PlanRegenerateError:
        pass

    return FileResponse(open(match["file_path"], "rb"), as_attachment=True, filename=filename)


# ---------- จัดการ Template (ดาวน์โหลด/อัปโหลด/ดูประวัติเวอร์ชัน/กู้คืน/ลบ) ----------

def template_list(request):
    templates = list_templates()
    return render(request, "cpall/template_list.html", {"templates": templates})


def template_download(request, key):
    if key not in TEMPLATE_REGISTRY:
        raise Http404
    path = TEMPLATE_REGISTRY[key]["path"]
    if not os.path.exists(path):
        raise Http404
    return FileResponse(open(path, "rb"), as_attachment=True, filename=os.path.basename(path))


def template_upload(request, key):
    if request.method != "POST":
        return redirect("cpall:template_list")
    if key not in TEMPLATE_REGISTRY:
        raise Http404

    is_htmx = request.headers.get("HX-Request") == "true"

    def error_response(message):
        if is_htmx:
            response = HttpResponse(status=400)
            response["HX-Trigger"] = json.dumps({"toast": {"message": message, "level": "error"}})
            return response
        return render(request, "cpall/template_upload_result.html",
                       {"success": False, "error": message, "key": key, "label": TEMPLATE_REGISTRY[key]["label"]})

    form = TemplateUploadForm(request.POST, request.FILES)
    if not form.is_valid():
        error_text = "; ".join(e for errs in form.errors.values() for e in errs)
        return error_response(error_text or "ยังไม่ได้เลือกไฟล์")
    new_file = form.cleaned_data["template_file"]

    os.makedirs(TEMP_UPLOAD_DIR, exist_ok=True)
    temp_path = os.path.join(TEMP_UPLOAD_DIR, f"{key}_{new_file.name}")
    with open(temp_path, "wb") as f:
        for chunk in new_file.chunks():
            f.write(chunk)

    try:
        upload_new_version(key, temp_path)
    except TemplateValidationError as e:
        # validate ไม่ผ่าน -> เวอร์ชัน/ไฟล์ live เดิมไม่ถูกแตะเลย ลบไฟล์ที่อัปโหลดมาทิ้ง
        if os.path.exists(temp_path):
            os.remove(temp_path)
        return error_response(str(e))

    # สำเร็จ -> ไปหน้าประวัติเวอร์ชันของ template นี้เลย (เห็นเวอร์ชันใหม่ + sku_count ที่ validate
    # ได้ อยู่ในบริบทจริง แทนที่จะเป็นหน้าสรุปผลโดดๆ)
    if is_htmx:
        response = HttpResponse(status=200)
        response["HX-Redirect"] = reverse("cpall:template_versions", args=[key])
        return response
    return redirect("cpall:template_versions", key=key)


def template_versions(request, key):
    """หน้าประวัติทุกเวอร์ชันของ template นี้ — กู้คืน/ลบได้จากหน้านี้"""
    if key not in TEMPLATE_REGISTRY:
        raise Http404
    versions = list_versions(key)
    return render(request, "cpall/template_versions.html", {
        "key": key, "label": TEMPLATE_REGISTRY[key]["label"], "versions": versions,
    })


def template_version_restore(request, key, version_id):
    if request.method != "POST":
        return redirect("cpall:template_versions", key=key)
    if key not in TEMPLATE_REGISTRY:
        raise Http404
    is_htmx = request.headers.get("HX-Request") == "true"
    try:
        version = restore_to_version(key, version_id)
        toast = {"message": f"ใช้เวอร์ชัน {version.version_number} แล้ว", "level": "success"}
    except TemplateValidationError as e:
        if not is_htmx:
            return redirect("cpall:template_versions", key=key)
        toast = {"message": str(e), "level": "error"}

    if is_htmx:
        versions = list_versions(key)
        response = render(request, "cpall/_template_version_list.html", {"key": key, "versions": versions})
        response["HX-Trigger"] = json.dumps({"toast": toast})
        return response
    return redirect("cpall:template_versions", key=key)


def template_version_delete(request, key, version_id):
    if request.method != "POST":
        return redirect("cpall:template_versions", key=key)
    if key not in TEMPLATE_REGISTRY:
        raise Http404
    is_htmx = request.headers.get("HX-Request") == "true"
    try:
        delete_version(key, version_id)
        toast = {"message": "ลบเวอร์ชันสำเร็จ", "level": "success"}
    except (TemplateInUseError, TemplateValidationError) as e:
        toast = {"message": str(e), "level": "error"}
        if not is_htmx:
            versions = list_versions(key)
            return render(request, "cpall/template_versions.html", {
                "key": key, "label": TEMPLATE_REGISTRY[key]["label"], "versions": versions, "error": str(e),
            })

    versions = list_versions(key)
    if is_htmx:
        response = render(request, "cpall/_template_version_list.html", {"key": key, "versions": versions})
        response["HX-Trigger"] = json.dumps({"toast": toast})
        return response
    return render(request, "cpall/template_versions.html", {
        "key": key, "label": TEMPLATE_REGISTRY[key]["label"], "versions": versions,
    })


# ---------- ดูตารางตัวเลขจริงในหน้าเว็บ (ไม่ต้องดาวน์โหลด Excel) ----------

def view_production_table(request, plan_run_id):
    detail = get_plan_run_detail(plan_run_id)
    if detail is None or detail["production_plan_status"] != "success":
        raise Http404
    # ลองอ่านจาก plan_sku_result ก่อน (ค่าจริงจาก LibreOffice) — ถ้าไม่มี (แผนเก่าก่อนมีระบบนี้ หรือ
    # extraction ตอนสร้างแผนล้มเหลว) ค่อย fallback ไปเปิดไฟล์ Excel + สูตรจำลอง Python แบบเดิม
    table = get_production_plan_table_from_db(plan_run_id)
    if not table["rows"]:
        table = get_production_plan_table(detail["production_plan_path"])
    # Django template ดึงค่าจาก dict ด้วย key ที่เป็นตัวแปรไม่ได้ตรงๆ -> จัดเป็น list ของ (qty, pack_text)
    # คู่กันตามลำดับคอลัมน์ไว้ก่อน เพื่อวนลูปแสดงคู่กันในเทมเพลตได้ง่ายๆ
    for row in table["rows"]:
        row["location_items"] = [
            {"qty": row["qty_by_location"].get(loc), "pack_text": row["pack_text_by_location"].get(loc)}
            for loc in table["sub_locations"]
        ]
    return render(request, "cpall/table_production.html", {"plan": detail, "table": table})


def view_logistic_table(request, plan_run_id, group_name):
    detail = get_plan_run_detail(plan_run_id)
    if detail is None:
        raise Http404
    match = next((lp for lp in detail["logistic_plans"]
                  if lp["group_name"] == group_name and lp["status"] == "success"), None)
    if match is None:
        raise Http404
    table = get_logistic_plan_table_from_db(plan_run_id, group_name)
    if not table["rows"]:
        table = get_logistic_plan_table(match["file_path"], group_name)
    for row in table["rows"]:
        row["location_items"] = [
            {"qty": row["qty_by_column"].get(col), "pack_text": row["pack_text_by_column"].get(col)}
            for col in table["columns"]
        ]
    return render(request, "cpall/table_logistic.html",
                  {"plan": detail, "table": table, "group_name": group_name})


# ---------- ลบ PO / ลบแผน ----------

def delete_po_import_view(request, po_import_id):
    if request.method != "POST":
        return redirect("cpall:po_list")
    is_htmx = request.headers.get("HX-Request") == "true"
    try:
        delete_po_import(po_import_id)
    except POInUseError as e:
        if is_htmx:
            # 409 = ไม่สำเร็จ -> HTMX จะไม่เอา response ไปแทนที่แถวเดิม (แถวยังอยู่ครบ) แค่โชว์ toast แดง
            response = HttpResponse(status=409)
            response["HX-Trigger"] = json.dumps({"toast": {"message": str(e), "level": "error"}})
            return response
        return render(request, "cpall/plan_error.html", {"error": str(e)})

    if is_htmx:
        # ตัวเปล่า status 200 -> hx-swap="outerHTML" เอาไปแทนที่ <tr> เดิม = แถวหายไปจากตารางทันที
        response = HttpResponse(status=200)
        response["HX-Trigger"] = json.dumps({"toast": {"message": "ลบ PO สำเร็จ", "level": "success"}})
        return response
    return redirect("cpall:po_list")


def delete_plan_run_view(request, plan_run_id):
    if request.method != "POST":
        return redirect("cpall:plan_list")
    delete_plan_run(plan_run_id)
    if request.headers.get("HX-Request") == "true":
        response = HttpResponse(status=200)
        response["HX-Trigger"] = json.dumps({"toast": {"message": "ลบแผนสำเร็จ", "level": "success"}})
        return response
    return redirect("cpall:plan_list")


def template_view(request, key):
    if key not in TEMPLATE_REGISTRY:
        raise Http404
    sheet_name = request.GET.get("sheet")
    try:
        grid = get_template_grid(key, sheet_name=sheet_name)
    except TemplateValidationError as e:
        raise Http404(str(e))
    return render(request, "cpall/template_view.html", {"key": key, "grid": grid})
