from django.contrib import admin

from .models import LocationMapping, LogisticGroup, PoRequiredColumn, ProductionPlanConfig, ProductMaster, Vehicle


@admin.register(ProductMaster)
class ProductMasterAdmin(admin.ModelAdmin):
    list_display = ("barcode", "name_th", "pack_size", "unit_price", "is_active", "updated_at")
    list_editable = ("is_active",)
    list_filter = ("is_active",)
    search_fields = ("barcode", "name_th", "name_en")


@admin.register(LocationMapping)
class LocationMappingAdmin(admin.ModelAdmin):
    list_display = ("fc_code", "name_th", "group", "sub_location", "is_active", "updated_at")
    list_editable = ("is_active",)
    list_filter = ("group", "is_active")
    search_fields = ("fc_code", "name_th", "sub_location")


@admin.register(LogisticGroup)
class LogisticGroupAdmin(admin.ModelAdmin):
    """
    จัดการกลุ่มพื้นที่ของ Logistic Plan ผ่านหน้าเว็บโดยตรง (เดิม hardcode ในโค้ด ต้องแก้+deploy ใหม่
    ถ้าจะเพิ่มกลุ่มที่ 5 ขึ้นไป) — เพิ่มกลุ่มใหม่ที่นี่แล้วไปอัปโหลดไฟล์เทมเพลตของกลุ่มนั้นที่หน้า
    "Template" ในเว็บหลัก (จะขึ้นให้เลือกอัตโนมัติตาม template_key ที่ตั้งไว้ตรงนี้ ไม่ต้องแก้โค้ดเลย)
    """
    list_display = ("group_name", "template_key", "sheet_name", "display_order", "is_active")
    list_editable = ("display_order", "is_active")
    search_fields = ("group_name", "template_key")
    ordering = ("display_order", "group_name")

    def get_form(self, request, obj=None, **kwargs):
        form = super().get_form(request, obj, **kwargs)
        form.base_fields["template_key"].help_text = (
            "ต้องขึ้นต้นด้วย \"logistic_\" เสมอ เช่น \"logistic_ระยอง\" — บันทึกแล้วไปอัปโหลดไฟล์"
            "เทมเพลตของกลุ่มนี้ที่หน้า \"Template\" ในเว็บหลัก (จะขึ้นให้เลือกอัตโนมัติ)"
        )
        form.base_fields["sheet_name"].help_text = (
            "ชื่อ Sheet ในไฟล์เทมเพลตที่มีข้อมูลจริงของกลุ่มนี้ (ต้องตรงเป๊ะกับที่อยู่ในไฟล์ที่จะอัปโหลด)"
        )
        return form

    def get_exclude(self, request, obj=None):
        # ระบบตอนนี้มีลูกค้าเดียว (cpall) — ไม่ต้องให้ Admin เลือก customer เองทุกครั้งที่เพิ่มกลุ่ม
        # ใหม่ (ไม่มีประโยชน์ มีตัวเลือกเดียวอยู่แล้ว) ซ่อน field นี้แล้ว auto-fill ให้แทน (ดู save_model)
        return ("customer",)

    def save_model(self, request, obj, form, change):
        if not obj.customer_id:
            from customers.cpall.logic.db import get_cpall_customer_id
            obj.customer_id = get_cpall_customer_id()
        super().save_model(request, obj, form, change)


@admin.register(Vehicle)
class VehicleAdmin(admin.ModelAdmin):
    """รายชื่อรถที่มีจริง — Admin เพิ่ม/ปิดใช้งานเองได้ (2025-09-12) ไม่ผูกกับกลุ่มพื้นที่ใดๆ เพราะ
    รถคันเดียวใช้วิ่งกลุ่มไหนก็ได้ เลือกอิสระตอนดูแผนแต่ละครั้ง"""
    list_display = ("plate_number", "vehicle_size", "basket_capacity", "is_active")
    list_editable = ("is_active",)
    list_filter = ("vehicle_size", "is_active")
    search_fields = ("plate_number",)
    ordering = ("basket_capacity",)

    def get_exclude(self, request, obj=None):
        return ("customer",)

    def save_model(self, request, obj, form, change):
        if not obj.customer_id:
            from customers.cpall.logic.db import get_cpall_customer_id
            obj.customer_id = get_cpall_customer_id()
        super().save_model(request, obj, form, change)


@admin.register(PoRequiredColumn)
class PoRequiredColumnAdmin(admin.ModelAdmin):
    """
    คอลัมน์ที่ต้องมีในไฟล์ PO Export จาก CP All — ถ้า CP All เปลี่ยนชื่อคอลัมน์ในไฟล์ export ของเขา
    Admin แก้ที่นี่ได้เลย ไม่ต้องแตะไฟล์/โค้ดเลย (เดิม hardcode ในโค้ด แล้วย้ายไปไฟล์ YAML ตามลำดับ)
    """
    list_display = ("column_name", "display_order")
    list_editable = ("display_order",)
    search_fields = ("column_name",)
    ordering = ("display_order", "column_name")

    def get_form(self, request, obj=None, **kwargs):
        form = super().get_form(request, obj, **kwargs)
        form.base_fields["column_name"].help_text = (
            "พิมพ์ชื่อคอลัมน์ตรงๆ ธรรมดา ไม่ต้องเผื่อช่องว่างหัว/ท้ายเอง (ระบบเทียบแบบไม่สนใจช่องว่าง"
            "หัว/ท้ายอยู่แล้ว แม้ไฟล์ต้นฉบับจริงจะมีช่องว่างต่อท้ายบางชื่อคอลัมน์ก็ตาม)"
        )
        return form

    def get_exclude(self, request, obj=None):
        return ("customer",)

    def save_model(self, request, obj, form, change):
        if not obj.customer_id:
            from customers.cpall.logic.db import get_cpall_customer_id
            obj.customer_id = get_cpall_customer_id()
        super().save_model(request, obj, form, change)


@admin.register(ProductionPlanConfig)
class ProductionPlanConfigAdmin(admin.ModelAdmin):
    """ตั้งค่า Production Plan (ตอนนี้มีแค่ชื่อ sheet) — แก้ที่นี่ถ้าเทมเพลตเปลี่ยนชื่อ sheet"""
    list_display = ("sheet_name",)

    def get_exclude(self, request, obj=None):
        return ("customer",)

    def save_model(self, request, obj, form, change):
        if not obj.customer_id:
            from customers.cpall.logic.db import get_cpall_customer_id
            obj.customer_id = get_cpall_customer_id()
        super().save_model(request, obj, form, change)

    def has_add_permission(self, request):
        # มีแค่แถวเดียวเสมอ (ต่อลูกค้า) — ไม่ให้เพิ่มซ้ำ กันสับสนว่าใช้แถวไหน
        return not ProductionPlanConfig.objects.exists()
