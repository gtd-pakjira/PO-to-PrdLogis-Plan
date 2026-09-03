"""
portal/views.py — หน้าแรกสุด: ภาพรวมทุกลูกค้า (สถานะล่าสุด + ปุ่มเข้าแต่ละลูกค้า)

ตอนนี้มีแค่ลูกค้าเดียว (cpall) แต่โครงนี้พร้อมรองรับลูกค้าเพิ่มในอนาคต — เพิ่มลูกค้าใหม่ผ่าน
Django Admin (/admin/) แล้ว list_customers() นี้จะเห็นเองอัตโนมัติ ไม่ต้องแก้โค้ดหน้านี้เลย
"""
from django.shortcuts import render

from core.models import Customer
from customers.cpall.logic.plan_runner import list_plan_runs
from customers.cpall.logic.po_parser import list_po_imports


def index(request):
    customers = []
    for c in Customer.objects.all().order_by("name_th"):
        # ตอนนี้มีแค่ cpall โมดูลเดียว — เช็คสถานะล่าสุดได้ตรงๆ แบบนี้ไปก่อน
        # (พอมีลูกค้าเจ้าที่ 2 ต้องมีจุดต่อ "code -> โมดูลไหนดึงสถานะยังไง" แยกให้ถูกต้อง)
        latest_po = None
        latest_plan = None
        if c.code == "cpall":
            po_imports = list_po_imports(limit=1)
            latest_po = po_imports[0] if po_imports else None
            plan_runs = list_plan_runs(limit=1)
            latest_plan = plan_runs[0] if plan_runs else None

        customers.append({
            "customer": c,
            "latest_po": latest_po,
            "latest_plan": latest_plan,
            "url": f"/{c.code}/" if c.code == "cpall" else None,  # โมดูลอื่นยังไม่มี URL จริง
        })

    return render(request, "portal/index.html", {"customers": customers})
