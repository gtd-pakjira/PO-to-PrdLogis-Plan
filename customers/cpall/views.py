"""
views.py — หน้าเว็บทั้งหมดของโมดูล cpall (CP All / 7-11)
"""
import io
import json
import os
import zipfile
from datetime import date, datetime
from urllib.parse import quote

from django.http import FileResponse, Http404, HttpResponse
from django.shortcuts import redirect, render
from django.urls import reverse

from customers.cpall.forms import ImportPOForm, TemplateUploadForm
from customers.cpall.logic.grouping import InactiveSkuOrderedError, ReconciliationError
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
    get_skipped_skus,
)
from customers.cpall.logic.po_parser import (
    POInUseError,
    POParseError,
    check_duplicate_rows,
    check_unknown_locations,
    check_unknown_skus,
    delete_po_import,
    list_po_imports,
    list_po_imports_paginated,
    load_po_to_db,
)
from customers.cpall.logic.po_regenerator import PORegenerateError, regenerate_po_file_bytes
from customers.cpall.logic.po_view_data import get_po_detail
from customers.cpall.logic.product_master_manager import save_product
from customers.cpall.logic.template_manager import (
    TemplateInUseError,
    TemplateValidationError,
    delete_version,
    get_template_grid,
    get_template_registry,
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
        duplicate_groups = check_duplicate_rows(saved_path)
    except POParseError as e:
        return error_response(f"ไฟล์ PO มีปัญหา: {e}")
    except Exception as e:
        return error_response(f"อ่านไฟล์ล้มเหลว: {type(e).__name__}: {e}", status=500)

    if duplicate_groups:
        # เจอรายการที่อาจซ้ำ (po_number+fc_code+barcode+line_no ตรงกันเป๊ะ) — ไม่ import ต่อทันที
        # ให้ Admin เห็นรายละเอียดแล้วเลือกเองว่า Continue (import ทุกแถวตามไฟล์จริง ไม่ตัดอะไรออก)
        # หรือ Stop (ยกเลิก ลบไฟล์ที่ค้างทิ้ง) — เก็บ path+วันที่ไว้ใน session (ไม่ใช่ hidden field ใน
        # form กัน path ถูกแก้ผ่าน browser dev tools ได้) รอ Admin ตัดสินใจที่หน้า confirm_duplicates
        request.session["pending_po_import"] = {
            "saved_path": saved_path,
            "production_date": production_date.isoformat(),
            "po_date": po_date.isoformat(),
            "original_filename": po_file.name,
        }
        context = {"duplicate_groups": duplicate_groups, "original_filename": po_file.name}
        if is_htmx:
            response = HttpResponse(status=200)
            response["HX-Redirect"] = reverse("cpall:confirm_duplicates")
            return response
        return render(request, "cpall/confirm_duplicates.html", context)

    try:
        po_import_id = load_po_to_db(saved_path, production_date, po_date, imported_by="web")
    except POParseError as e:
        return error_response(f"ไฟล์ PO มีปัญหา: {e}")
    except Exception as e:
        return error_response(f"นำเข้าล้มเหลว: {type(e).__name__}: {e}", status=500)

    # นำเข้าสำเร็จ -> ข้อมูลทุกคอลัมน์ของไฟล์ (ไม่ใช่แค่ 12 คอลัมน์ที่ระบบใช้คำนวณ) ถูกเก็บไว้ครบใน
    # po_import.column_order + po_line.all_values แล้ว (ดู po_parser.py's parse_po_file) ไม่จำเป็นต้อง
    # เก็บไฟล์ต้นฉบับไว้บนดิสก์อีกต่อไปเลย — ดาวน์โหลดย้อนหลังทีหลัง จะสร้างไฟล์ใหม่จากข้อมูลใน DB แทน
    # (ดู po_regenerator.py) ทดสอบยืนยันแล้วว่าตรงกับต้นฉบับทุกเซลล์ 100%
    if os.path.exists(saved_path):
        os.remove(saved_path)

    unknown_locations = check_unknown_locations(po_import_id)
    if unknown_locations:
        # ไม่ใช่ dead-end แล้ว — พาไปหน้าเลือก mapping เลย (ข้อมูล PO import สำเร็จแล้วจริงๆ ใน DB
        # แค่ยังมีรหัสสถานที่ที่ไม่รู้จักกลุ่มพื้นที่ ให้ Admin เลือกตรงนี้ได้เลย)
        if is_htmx:
            response = HttpResponse(status=200)
            response["HX-Redirect"] = reverse("cpall:resolve_locations", args=[po_import_id])
            return response
        return redirect("cpall:resolve_locations", po_import_id=po_import_id)

    unknown_skus = check_unknown_skus(po_import_id)
    if unknown_skus:
        # ต่างจาก location ตรงที่ไม่บังคับ (product_master ไม่มีผลต่อการคำนวณเลย) แต่ยังพาไปหน้านี้
        # เพื่อ "แนะนำ" ให้กรอกไว้ให้ครบ (มีปุ่มข้ามในหน้านั้นให้กดผ่านได้เลยถ้าไม่อยากกรอกตอนนี้)
        if is_htmx:
            response = HttpResponse(status=200)
            response["HX-Redirect"] = reverse("cpall:resolve_products", args=[po_import_id])
            return response
        return redirect("cpall:resolve_products", po_import_id=po_import_id)

    # สำเร็จสมบูรณ์ -> ไปหน้า PO ทั้งหมดเลย (เห็นผลลัพธ์อยู่ในบริบทของรายการจริง แทนที่จะเป็นหน้า
    # สรุปผลโดดๆ ที่ต้องกดออกไปอีกที)
    if is_htmx:
        response = HttpResponse(status=200)
        response["HX-Redirect"] = reverse("cpall:po_list")
        return response
    return redirect("cpall:po_list")


def confirm_duplicates(request):
    """
    หน้า "พบรายการที่อาจซ้ำ" — Admin เลือก Continue (import ทุกแถวตามไฟล์จริง ไม่ตัดอะไรออกเลย)
    หรือ Stop (ยกเลิก ลบไฟล์ที่ค้างทิ้ง) — path/วันที่ที่ต้องใช้เก็บไว้ใน session จาก import_submit()
    ไม่รับผ่าน POST body เอง (กัน path ถูกแก้ผ่าน browser dev tools ได้)
    """
    pending = request.session.get("pending_po_import")
    if pending is None:
        return render(request, "cpall/plan_error.html",
                       {"error": "ไม่พบไฟล์ที่รออยู่ (session อาจหมดอายุ) กรุณาอัปโหลดใหม่"})

    is_htmx = request.headers.get("HX-Request") == "true"

    if request.method == "POST":
        action = request.POST.get("action")
        del request.session["pending_po_import"]  # ใช้ครั้งเดียวจบ ไม่ว่าจะกดปุ่มไหน

        if action == "stop":
            if os.path.exists(pending["saved_path"]):
                os.remove(pending["saved_path"])
            if is_htmx:
                response = HttpResponse(status=200)
                response["HX-Trigger"] = json.dumps(
                    {"toast": {"message": "ยกเลิกการนำเข้าแล้ว", "level": "success"}})
                response["HX-Redirect"] = reverse("cpall:import_form")
                return response
            return redirect("cpall:import_form")

        # action == "continue" (หรือค่าอื่นที่ไม่ใช่ stop ก็ถือว่า continue ไปเลย ปลอดภัยกว่าปฏิเสธ
        # การนำเข้าไปเฉยๆ โดยไม่มีเหตุผล) — import ทุกแถวจริง ไม่ตัดอะไรออกแม้จะซ้ำก็ตาม
        production_date = date.fromisoformat(pending["production_date"])
        po_date = date.fromisoformat(pending["po_date"])
        try:
            po_import_id = load_po_to_db(pending["saved_path"], production_date, po_date, imported_by="web")
        except POParseError as e:
            return render(request, "cpall/plan_error.html", {"error": f"ไฟล์ PO มีปัญหา: {e}"})
        except Exception as e:
            return render(request, "cpall/plan_error.html",
                           {"error": f"นำเข้าล้มเหลว: {type(e).__name__}: {e}"})

        if os.path.exists(pending["saved_path"]):
            os.remove(pending["saved_path"])

        # ต่อ flow เดิมเป๊ะ (เหมือน import_submit ตอนไม่มี duplicate) — เช็ค location/SKU ที่ไม่รู้จัก
        unknown_locations = check_unknown_locations(po_import_id)
        if unknown_locations:
            if is_htmx:
                response = HttpResponse(status=200)
                response["HX-Redirect"] = reverse("cpall:resolve_locations", args=[po_import_id])
                return response
            return redirect("cpall:resolve_locations", po_import_id=po_import_id)

        unknown_skus = check_unknown_skus(po_import_id)
        if unknown_skus:
            if is_htmx:
                response = HttpResponse(status=200)
                response["HX-Redirect"] = reverse("cpall:resolve_products", args=[po_import_id])
                return response
            return redirect("cpall:resolve_products", po_import_id=po_import_id)

        if is_htmx:
            response = HttpResponse(status=200)
            response["HX-Redirect"] = reverse("cpall:po_list")
            return response
        return redirect("cpall:po_list")

    # GET — แสดงหน้ายืนยัน (เรียก check_duplicate_rows ใหม่อีกครั้งจากไฟล์ที่ยังค้างอยู่ เผื่อ Admin
    # รีเฟรชหน้านี้ — ไม่ query จาก session เพราะ session เก็บแค่ path ไม่ได้เก็บรายละเอียดกลุ่มที่ซ้ำ)
    try:
        duplicate_groups = check_duplicate_rows(pending["saved_path"])
    except Exception:
        duplicate_groups = []
    return render(request, "cpall/confirm_duplicates.html", {
        "duplicate_groups": duplicate_groups, "original_filename": pending["original_filename"],
    })


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

        # location ครบหมดแล้ว -> เช็คต่อว่ามีสินค้าที่ยังไม่รู้จักไหม (ไม่บังคับ แค่แนะนำ)
        unknown_skus = check_unknown_skus(po_import_id)
        if unknown_skus:
            if is_htmx:
                response = HttpResponse(status=200)
                response["HX-Redirect"] = reverse("cpall:resolve_products", args=[po_import_id])
                return response
            return redirect("cpall:resolve_products", po_import_id=po_import_id)

        if is_htmx:
            response = HttpResponse(status=200)
            response["HX-Redirect"] = reverse("cpall:po_list")
            return response
        return redirect("cpall:po_list")

    return render(request, "cpall/resolve_locations.html", {
        "po_import_id": po_import_id, "unknown_locations": unknown_locations,
        "existing_groups": get_existing_groups(),
    })


def resolve_products(request, po_import_id):
    """ให้ Admin กรอกข้อมูลสินค้าที่ยังไม่รู้จักผ่านเว็บ — ต่างจาก resolve_locations ตรงที่ไม่บังคับ
    (product_master ไม่มีผลต่อการคำนวณเลย มีไว้แค่แสดงชื่อสินค้าที่หน้ากรอกยอดเผื่อ) มีปุ่มข้ามได้เสมอ"""
    unknown_skus = check_unknown_skus(po_import_id)
    is_htmx = request.headers.get("HX-Request") == "true"

    if request.method == "POST":
        for barcode, item_name, net_case_price in unknown_skus:
            name_th = request.POST.get(f"name_th_{barcode}", "").strip() or item_name
            name_en = request.POST.get(f"name_en_{barcode}", "").strip()
            pack_size_str = request.POST.get(f"pack_size_{barcode}", "").strip()
            unit_price_str = request.POST.get(f"unit_price_{barcode}", "").strip()

            if not pack_size_str:
                continue  # ยังไม่ได้กรอกอันนี้ -> ข้ามไปก่อน (ไม่บังคับ ต่างจาก location)
            try:
                pack_size = int(pack_size_str)
                # ราคาไม่ได้กรอกเอง -> ใช้ราคาจาก PO ที่ auto-fill ไว้ให้แล้ว (ถ้ามี) เป็น fallback
                unit_price = float(unit_price_str) if unit_price_str else net_case_price
            except ValueError:
                continue
            save_product(barcode, name_th, name_en or None, pack_size, unit_price)

        remaining = check_unknown_skus(po_import_id)
        if remaining:
            context = {"po_import_id": po_import_id, "unknown_skus": remaining}
            if is_htmx:
                return render(request, "cpall/_resolve_products_form.html", context)
            return render(request, "cpall/resolve_products.html", context)

        if is_htmx:
            response = HttpResponse(status=200)
            response["HX-Redirect"] = reverse("cpall:po_list")
            return response
        return redirect("cpall:po_list")

    return render(request, "cpall/resolve_products.html", {
        "po_import_id": po_import_id, "unknown_skus": unknown_skus,
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

    # ไปหน้ากรอกยอดเผื่อเสมอ — ขึ้นทุกครั้งที่สร้างแผน ไม่ใช่แค่ตอนมีรอบเช้าต่างจังหวัด
    # (ยอดเผื่อเป็นข้อมูลสำคัญที่ Admin ต้องยืนยันทุกรอบ ไม่ใช่แค่กลุ่มเช้าต่างจังหวัด)
    if is_htmx:
        response = HttpResponse(status=200)
        response["HX-Redirect"] = _buffer_form_url(po_import_ids)
        return response
    return redirect_to_buffer_form(request, po_import_ids)


def _buffer_form_url(po_import_ids):
    from urllib.parse import urlencode
    qs = urlencode({"po_import_ids": ",".join(str(x) for x in po_import_ids)})
    return f"{reverse('cpall:buffer_form')}?{qs}"


def redirect_to_buffer_form(request, po_import_ids):
    return redirect(_buffer_form_url(po_import_ids))


def buffer_form(request):
    import openpyxl

    from customers.cpall.logic.excel_export import (
        SHEET_NAME as PP_SHEET_NAME,
    )
    from customers.cpall.logic.excel_export import TEMPLATE_PATH as PP_TEMPLATE_PATH
    from customers.cpall.logic.excel_export import (
        _find_sku_header_rows as _find_pp_sku_header_rows,
    )
    from customers.cpall.models import PlanSkuResult, ProductMaster

    po_import_ids_str = request.GET.get("po_import_ids", "")
    po_import_ids = [int(x) for x in po_import_ids_str.split(",") if x]
    if not po_import_ids:
        return render(request, "cpall/plan_error.html", {"error": "ไม่พบรอบ PO ที่เลือกไว้"})

    # ดึง SKU จาก Production Plan template (ครบ 19 ตัว เรียงตามลำดับในแผนผลิต) —
    # เดิมดึงจากเทมเพลตรอบเช้าต่างจังหวัด (18 ตัว) ซึ่งขาดพุทราจีนและไม่ตรงกับ Production Plan
    wb = openpyxl.load_workbook(PP_TEMPLATE_PATH)
    ws = wb[PP_SHEET_NAME]
    header_rows = _find_pp_sku_header_rows(ws)
    # เรียงตามตำแหน่งแถวในไฟล์ (ตามลำดับใน Production Plan จริง)
    barcodes = [bc for bc, _ in sorted(header_rows.items(), key=lambda kv: kv[1])]

    # default ยอดเผื่อ: ดึงจาก buffer_qty ล่าสุดที่เคยบันทึกไว้ใน plan_sku_result —
    # ไม่ใช่จากไฟล์เทมเพลต (ซึ่งเป็นค่าเก่าที่ Admin กรอกไว้ครั้งแรก ไม่ใช่ล่าสุด)
    default_buffer = {}
    last_plan = PlanSkuResult.objects.filter(
        sheet_type="production", buffer_qty__isnull=False
    ).order_by("-id").first()
    if last_plan:
        for row in PlanSkuResult.objects.filter(
            plan_run_id=last_plan.plan_run_id, sheet_type="production", buffer_qty__isnull=False
        ).values("barcode", "buffer_qty").distinct("barcode"):
            if row["buffer_qty"] is not None:
                default_buffer[row["barcode"]] = float(row["buffer_qty"])

    name_lookup = {s.barcode: s.name_th for s in ProductMaster.objects.filter(barcode__in=barcodes)}
    sku_rows = [
        {"barcode": bc, "name_th": name_lookup.get(bc, bc), "default_buffer": default_buffer.get(bc, 0)}
        for bc in barcodes
    ]

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
    except InactiveSkuOrderedError as e:
        return error_response(str(e), status=409)
    except ReconciliationError as e:
        return error_response(str(e), status=409)
    except Exception as e:
        return error_response(f"สร้างแผนล้มเหลว: {type(e).__name__}: {e}", status=500)

    if is_htmx:
        # ใช้ "replaceLocation" (custom event ที่ตัวเองทำ window.location.replace()) แทน HX-Redirect
        # ธรรมดา — HX-Redirect จะ "เพิ่ม" หน้าแผนใหม่เข้า browser history (window.location = url) ทำให้
        # หน้ากรอกยอดเผื่อ (ที่ควรเป็นแค่ "ขั้นตอนแวะผ่าน" ไม่ใช่ปลายทาง) ยังค้างอยู่ใน history —
        # พอกดปุ่มย้อนกลับจากหน้าแผน จะเด้งไปหน้ากรอกยอดเผื่อแทนที่จะเป็นหน้า PO ที่กดสร้างแผนมาจริงๆ
        # .replace() แทนที่ entry ปัจจุบัน (หน้ากรอกยอดเผื่อ) เลย ทำให้กด back ข้ามไปหน้า PO ตรงๆ
        response = HttpResponse(status=200)
        response["HX-Trigger"] = json.dumps({
            "replaceLocation": {"url": reverse("cpall:view_plan", args=[result["plan_run_id"]])}
        })
        return response
    return redirect("cpall:view_plan", plan_run_id=result["plan_run_id"])


def _set_download_filename(response, filename):
    """
    ตั้งชื่อไฟล์ดาวน์โหลดแบบรองรับภาษาไทยให้ถูกต้องตามมาตรฐาน RFC 6266 (filename*=UTF-8''...)
    ไม่ใช้ resp["Content-Disposition"] = f'...filename="{filename}"...' ตรงๆ เพราะ Django จะ auto-encode
    ทั้ง header เป็น MIME encoded-word แบบเก่า (=?utf-8?b?...?=) ทันทีที่เจอตัวอักษรที่ไม่ใช่ ASCII —
    ซึ่งเป็นมาตรฐานสำหรับ email header ไม่ใช่ HTTP header บางเบราว์เซอร์/ตัวจัดการดาวน์โหลด parse ผิด
    แล้ว fallback ไปใช้ชื่อ "download" เฉยๆ (ตรงกับปัญหาที่เจอจริง) — เข้ารหัสเองแบบ RFC 6266 แทน พร้อม
    ชื่อสำรอง ASCII ควบคู่กันไปด้วย เผื่อเบราว์เซอร์เก่ามากๆ ไม่รู้จัก filename*= เลย (นามสกุลของชื่อ
    สำรองต้องตรงกับไฟล์จริงเสมอ — ไม่ hardcode .xlsx เพราะไฟล์นี้อาจเป็น .zip ก็ได้)
    """
    ext = os.path.splitext(filename)[1] or ".xlsx"
    encoded = quote(filename)
    response["Content-Disposition"] = f"attachment; filename=\"download{ext}\"; filename*=UTF-8''{encoded}"


def view_plan(request, plan_run_id):
    detail = get_plan_run_detail(plan_run_id)
    if detail is None:
        return render(request, "cpall/plan_not_found.html", {"plan_run_id": plan_run_id}, status=404)
    plan_name = PlanRun.objects.get(id=plan_run_id).get_short_label()
    skipped_skus = get_skipped_skus(plan_run_id)
    return render(request, "cpall/plan_view.html", {
        "plan": detail, "plan_name": plan_name, "skipped_skus": skipped_skus,
    })


def download_production(request, plan_run_id):
    detail = get_plan_run_detail(plan_run_id)
    if detail is None or not detail["production_plan_path"]:
        raise Http404
    filename = f"{PlanRun.objects.get(id=plan_run_id).get_display_name(prefix='แพลน')}.xlsx"

    # ลองสร้างไฟล์ใหม่จากข้อมูลดิบก่อนเสมอ (มีสูตรจริงครบ ตรงกับที่ตกลงกันไว้) — ถ้าทำไม่ได้ (แผนเก่า
    # ก่อนมีระบบ data-first) fallback ไปเสิร์ฟไฟล์ที่เก็บไว้บนดิสก์แบบเดิม (ยังไม่ได้ลบไฟล์เก่าทิ้งในเฟสนี้)
    try:
        content = regenerate_production_plan_bytes(plan_run_id)
        response = HttpResponse(
            content,
            content_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
        _set_download_filename(response, filename)
        return response
    except PlanRegenerateError:
        pass

    response = FileResponse(open(detail["production_plan_path"], "rb"),
                             content_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    _set_download_filename(response, filename)
    return response


def download_logistic(request, plan_run_id, group_name):
    detail = get_plan_run_detail(plan_run_id)
    if detail is None:
        raise Http404
    match = next((lp for lp in detail["logistic_plans"] if lp["group_name"] == group_name), None)
    if match is None or not match["file_path"]:
        raise Http404
    filename = f"{PlanRun.objects.get(id=plan_run_id).get_display_name(prefix=group_name)}.xlsx"

    try:
        content = regenerate_logistic_plan_bytes(plan_run_id, group_name)
        response = HttpResponse(
            content,
            content_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
        _set_download_filename(response, filename)
        return response
    except PlanRegenerateError:
        pass

    response = FileResponse(open(match["file_path"], "rb"),
                             content_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    _set_download_filename(response, filename)
    return response


def view_po_detail(request, po_import_id):
    detail = get_po_detail(po_import_id)
    if detail is None:
        raise Http404
    return render(request, "cpall/po_detail.html", {"po": detail})


def download_po(request, po_import_id):
    """สร้างไฟล์ PO ใหม่จากข้อมูลใน database (ครบทุกคอลัมน์เหมือนต้นฉบับ ทดสอบยืนยันตรงกันทุกเซลล์)
    — ไม่ได้เก็บไฟล์ต้นฉบับไว้ถาวรอีกต่อไป แต่ถ้าเป็น PO เก่าที่ยังมีไฟล์ค้างอยู่จากก่อนเปลี่ยนมาใช้
    ระบบนี้ ให้เสิร์ฟไฟล์เดิมนั้นตรงๆ ก่อน (ของจริงย่อมดีกว่าของสร้างใหม่เสมอถ้ามีอยู่จริง)"""
    detail = get_po_detail(po_import_id)
    if detail is None:
        raise Http404

    source_path = detail["source_filename"]
    if source_path and os.path.exists(source_path):
        response = FileResponse(
            open(source_path, "rb"),
            content_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
        _set_download_filename(response, detail["display_filename"])
        return response

    try:
        content = regenerate_po_file_bytes(po_import_id)
    except PORegenerateError:
        raise Http404
    response = HttpResponse(
        content, content_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )
    _set_download_filename(response, detail["display_filename"])
    return response


def download_all_zip(request, plan_run_id):
    """โหลดทั้งแผน (Production Plan + Logistic Plan ทุกกลุ่มที่สำเร็จ) รวมเป็นไฟล์ ZIP เดียว —
    ใช้ตัวสร้างไฟล์เดียวกับการโหลดทีละไฟล์ทุกอย่าง (มีสูตรจริงครบเหมือนกัน) แค่รวมเข้า zip ในหน่วยความจำ
    ไม่ผ่านไฟล์ชั่วคราวบนดิสก์เลย (เหมือน pattern ที่ใช้กับการโหลดทีละไฟล์)"""
    detail = get_plan_run_detail(plan_run_id)
    if detail is None:
        raise Http404
    plan = PlanRun.objects.get(id=plan_run_id)

    buffer = io.BytesIO()
    with zipfile.ZipFile(buffer, "w", zipfile.ZIP_DEFLATED) as zf:
        if detail["production_plan_status"] == "success":
            try:
                content = regenerate_production_plan_bytes(plan_run_id)
            except PlanRegenerateError:
                content = None
                if detail["production_plan_path"] and os.path.exists(detail["production_plan_path"]):
                    with open(detail["production_plan_path"], "rb") as f:
                        content = f.read()
            if content:
                zf.writestr(f"{plan.get_display_name(prefix='แพลน')}.xlsx", content)

        for lp in detail["logistic_plans"]:
            if lp["status"] != "success":
                continue
            try:
                content = regenerate_logistic_plan_bytes(plan_run_id, lp["group_name"])
            except PlanRegenerateError:
                content = None
                if lp["file_path"] and os.path.exists(lp["file_path"]):
                    with open(lp["file_path"], "rb") as f:
                        content = f.read()
            if content:
                zf.writestr(f"{plan.get_display_name(prefix=lp['group_name'])}.xlsx", content)

    buffer.seek(0)
    response = HttpResponse(buffer.getvalue(), content_type="application/zip")
    zip_filename = f"{plan.get_display_name(prefix='แพลนทั้งหมด')}.zip"
    _set_download_filename(response, zip_filename)
    return response


# ---------- จัดการ Template (ดาวน์โหลด/อัปโหลด/ดูประวัติเวอร์ชัน/กู้คืน/ลบ) ----------

def template_list(request):
    templates = list_templates()
    return render(request, "cpall/template_list.html", {"templates": templates})


def template_download(request, key):
    registry = get_template_registry()
    if key not in registry:
        raise Http404
    path = registry[key]["path"]
    if not os.path.exists(path):
        raise Http404
    return FileResponse(open(path, "rb"), as_attachment=True, filename=os.path.basename(path))


def template_upload(request, key):
    if request.method != "POST":
        return redirect("cpall:template_list")
    registry = get_template_registry()
    if key not in registry:
        raise Http404

    is_htmx = request.headers.get("HX-Request") == "true"

    def error_response(message):
        if is_htmx:
            response = HttpResponse(status=400)
            response["HX-Trigger"] = json.dumps({"toast": {"message": message, "level": "error"}})
            return response
        return render(request, "cpall/template_upload_result.html",
                       {"success": False, "error": message, "key": key, "label": registry[key]["label"]})

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
        upload_new_version(key, temp_path, original_filename=new_file.name)
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
    registry = get_template_registry()
    if key not in registry:
        raise Http404
    versions = list_versions(key)
    return render(request, "cpall/template_versions.html", {
        "key": key, "label": registry[key]["label"], "versions": versions,
    })


def template_version_restore(request, key, version_id):
    if request.method != "POST":
        return redirect("cpall:template_versions", key=key)
    if key not in get_template_registry():
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
    registry = get_template_registry()
    if key not in registry:
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
                "key": key, "label": registry[key]["label"], "versions": versions, "error": str(e),
            })

    versions = list_versions(key)
    if is_htmx:
        response = render(request, "cpall/_template_version_list.html", {"key": key, "versions": versions})
        response["HX-Trigger"] = json.dumps({"toast": toast})
        return response
    return render(request, "cpall/template_versions.html", {
        "key": key, "label": registry[key]["label"], "versions": versions,
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
    total_baskets = sum(row["basket_total"] or 0 for row in table["rows"])
    basket_total_by_column = table.get("basket_total_by_column", {})  # ไม่มีถ้าเป็น fallback
    # ไปใช้ get_logistic_plan_table() (แผนเก่ามากที่ไม่มีข้อมูลใน plan_sku_result เลย) — ปล่อยว่างไว้
    # ก็พอ (แสดงแค่ "รวมตะกร้าทั้งหมด" แบบเดิม ไม่มีตะกร้าต่อคอลัมน์ให้แผนเก่ากลุ่มนี้)

    # เลข PO จริงเบื้องหลัง "PO1"/"PO2" แต่ละคอลัมน์ — เอาไว้แสดง tooltip เฉยๆ ถ้าหาไม่ได้ (เช่น PO
    # ต้นทางถูกลบไปแล้ว) ไม่ให้กระทบหน้าตารางเลย แค่ไม่มี tooltip ให้
    from customers.cpall.logic.logistic_plan_export import get_po_number_by_column_label

    po_import_ids = [po["id"] for po in detail["po_imports"]]
    try:
        po_number_by_column = get_po_number_by_column_label(po_import_ids, group_name)
    except Exception:
        po_number_by_column = {}
    # เตรียมคู่ (คอลัมน์, เลข PO) ไว้ล่วงหน้า — Django template lookup แบบ dict.{{ loop_var }} หา
    # key ชื่อ "loop_var" ตรงๆ ไม่ resolve ค่าตัวแปรให้ ต้องจับคู่มาก่อนแบบนี้แทน
    columns_with_po = [(col, po_number_by_column.get(col)) for col in table["columns"]]
    # เหตุผลเดียวกัน — เตรียมคู่ (คอลัมน์, ตะกร้าต่อคอลัมน์) ไว้ล่วงหน้าด้วย
    columns_with_basket = [(col, basket_total_by_column.get(col)) for col in table["columns"]]

    return render(request, "cpall/table_logistic.html", {
        "plan": detail, "table": table, "group_name": group_name, "total_baskets": total_baskets,
        "columns_with_po": columns_with_po, "columns_with_basket": columns_with_basket,
    })


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
    if key not in get_template_registry():
        raise Http404
    sheet_name = request.GET.get("sheet")
    try:
        grid = get_template_grid(key, sheet_name=sheet_name)
    except TemplateValidationError as e:
        raise Http404(str(e))
    return render(request, "cpall/template_view.html", {"key": key, "grid": grid})
