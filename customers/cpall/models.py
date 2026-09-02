"""
cpall/models.py — Django Model ผูกกับตารางที่มีอยู่แล้ว (สร้าง/ดูแลโดย sql/schema.sql ไม่ใช่ Django
migration — managed = False ทุกตัว)

*** ข้อจำกัดที่ต้องรู้ ***
ในฐานข้อมูลจริง sku_master/location_mapping มี PK แบบ composite คือ (customer_id, barcode/fc_code)
เพื่อรองรับหลายลูกค้าใช้บาร์โค้ด/รหัสสถานที่ซ้ำกันได้ — แต่ Django Admin "ลงทะเบียน model ที่มี
composite primary key ไม่ได้" (ข้อจำกัดของ Django เอง) เลยต้องให้ barcode/fc_code เป็น Django PK
แบบเดี่ยวไปก่อน (เหมือนก่อนรื้อระบบ) ตอนนี้ใช้ได้ปกติเพราะมีลูกค้าเดียว (cpall) บาร์โค้ดไม่ชนกัน —
ถ้าวันหน้ามีลูกค้าที่ 2 ที่บาร์โค้ดซ้ำกับ cpall จริงๆ ต้องเปลี่ยนมาใช้ surrogate key (เพิ่มคอลัมน์ id
AutoField ในตาราง DB จริง แล้วผูก Django PK กับ id แทน) — ยังไม่ต้องทำตอนนี้เพราะยังไม่มีลูกค้าที่ 2
"""
from django.db import models

from core.models import Customer


class SkuMaster(models.Model):
    barcode = models.CharField(max_length=20, primary_key=True, verbose_name="บาร์โค้ด")
    customer = models.ForeignKey(Customer, on_delete=models.CASCADE, db_column="customer_id")
    name_th = models.TextField(verbose_name="ชื่อสินค้า (ไทย)")
    name_en = models.TextField(verbose_name="ชื่อสินค้า (อังกฤษ)", blank=True, null=True)
    pack_size = models.IntegerField(verbose_name="บรรจุ/ตก. (ชิ้นต่อลัง)")
    unit_price = models.DecimalField(max_digits=10, decimal_places=2, verbose_name="ราคา", blank=True, null=True)
    note = models.TextField(verbose_name="หมายเหตุ", blank=True, null=True)
    updated_at = models.DateTimeField(auto_now=True)

    class Meta:
        db_table = "sku_master"
        managed = False
        verbose_name = "SKU"
        verbose_name_plural = "SKU Master"

    def __str__(self):
        return f"{self.barcode} — {self.name_th}"


class LocationMapping(models.Model):
    fc_code = models.CharField(max_length=10, primary_key=True, verbose_name="รหัสสถานที่ (FC code)")
    customer = models.ForeignKey(Customer, on_delete=models.CASCADE, db_column="customer_id")
    name_th = models.TextField(verbose_name="ชื่อสถานที่")
    group = models.CharField(max_length=50, db_column="group", verbose_name="กลุ่มพื้นที่")
    sub_location = models.CharField(max_length=50, blank=True, null=True, verbose_name="จุดส่งย่อย")
    updated_at = models.DateTimeField(auto_now=True)

    class Meta:
        db_table = "location_mapping"
        managed = False
        verbose_name = "จุดส่ง"
        verbose_name_plural = "Location Mapping"

    def __str__(self):
        return f"{self.fc_code} — {self.name_th}"


class PoImport(models.Model):
    customer = models.ForeignKey(Customer, on_delete=models.CASCADE, db_column="customer_id")
    source_filename = models.TextField()
    imported_at = models.DateTimeField(auto_now_add=True)
    imported_by = models.CharField(max_length=100, blank=True, null=True)
    # ชื่อฟิลด์นี้เคยผิดความหมายมาตั้งแต่แรก (แก้ Phase 1) — production_date = วันที่ผลิต (เดิมเรียก
    # po_date ผิดๆ), po_date = วันที่ PO ตัวจริง (เดิมเรียก delivery_date ผิดๆ) ดูหัวไฟล์จริงของ
    # บางบัวทอง: "วันที่ผลิต ... ส่งวันที่ PO ..." — ค่าที่กรอกไว้เดิมถูกต้องอยู่แล้ว แก้แค่ชื่อ
    production_date = models.DateField(blank=True, null=True, verbose_name="วันที่ผลิต")
    po_date = models.DateField(blank=True, null=True, verbose_name="วันที่ PO")
    total_rows = models.IntegerField(blank=True, null=True)
    status = models.CharField(max_length=20, default="imported")

    class Meta:
        db_table = "po_import"
        managed = False
        verbose_name = "PO Import"
        verbose_name_plural = "PO Imports"

    def __str__(self):
        return f"#{self.id} — {self.source_filename}"


class PoLine(models.Model):
    po_import = models.ForeignKey(PoImport, on_delete=models.CASCADE, db_column="po_import_id", related_name="lines")
    po_number = models.CharField(max_length=30)
    po_date = models.DateField(blank=True, null=True)
    delivery_date = models.DateField(blank=True, null=True)
    delivery_time = models.CharField(max_length=10, blank=True, null=True)
    fc_code = models.CharField(max_length=10)
    delivery_location = models.TextField(blank=True, null=True)
    line_no = models.IntegerField(blank=True, null=True)
    barcode = models.CharField(max_length=20)
    item_name = models.TextField(blank=True, null=True)
    qty_ordered = models.DecimalField(max_digits=10, decimal_places=2)
    unit_type = models.CharField(max_length=10, blank=True, null=True)
    net_case_price = models.DecimalField(max_digits=10, decimal_places=2, blank=True, null=True)
    total_amount = models.DecimalField(max_digits=12, decimal_places=2, blank=True, null=True)

    class Meta:
        db_table = "po_line"
        managed = False
        verbose_name = "PO Line"
        verbose_name_plural = "PO Lines"

    def __str__(self):
        return f"{self.po_number} — {self.barcode}"


class TemplateVersion(models.Model):
    """
    เก็บทุกเวอร์ชันของไฟล์ Template ที่เคยอัปโหลดถาวร (Phase 1.6 sub-phase 1) — is_active=True มีได้
    แค่ 1 แถวต่อ (customer, template_key) เท่านั้น (บังคับผ่านโค้ดใน template_manager.py ไม่ใช่ DB
    constraint เพราะ Postgres ไม่มี "unique ตอน is_active=True" ตรงๆ ง่ายๆ)
    """
    customer = models.ForeignKey(Customer, on_delete=models.CASCADE, db_column="customer_id")
    template_key = models.CharField(max_length=50)
    version_number = models.IntegerField()
    file_path = models.TextField()
    is_active = models.BooleanField(default=False)
    uploaded_at = models.DateTimeField(auto_now_add=True)
    validation_summary = models.TextField(blank=True, null=True)

    class Meta:
        db_table = "template_version"
        managed = False
        verbose_name = "Template Version"
        verbose_name_plural = "Template Versions"
        ordering = ["-version_number"]

    def __str__(self):
        return f"{self.template_key} v{self.version_number}"


class PlanRun(models.Model):
    customer = models.ForeignKey(Customer, on_delete=models.CASCADE, db_column="customer_id")
    created_at = models.DateTimeField(auto_now_add=True)
    output_dir = models.TextField()
    production_plan_path = models.TextField(blank=True, null=True)
    production_plan_status = models.CharField(max_length=20, default="success")
    production_plan_error = models.TextField(blank=True, null=True)
    po_imports = models.ManyToManyField(PoImport, through="PlanRunImport", related_name="plan_runs")
    production_template_version = models.ForeignKey(
        TemplateVersion, on_delete=models.SET_NULL, db_column="production_template_version_id",
        blank=True, null=True, related_name="+",
    )

    class Meta:
        db_table = "plan_run"
        managed = False
        verbose_name = "Plan Run"
        verbose_name_plural = "Plan Runs"

    def __str__(self):
        return f"แผน #{self.id}"

    def get_display_name(self):
        """
        ชื่อแผนตามฟอร์แมตที่ตกลงกันไว้: แพลน_7-11_YYYY-MM-DD_HH:MM:SS:FF3
        คำนวณจาก created_at ทุกครั้ง ไม่ได้เก็บเป็นคอลัมน์แยก (created_at คือความจริงหนึ่งเดียวอยู่แล้ว)

        หมายเหตุ: created_at ที่ได้จาก ORM เป็น naive datetime (ไม่มี tzinfo) แต่ "ค่าจริงเป็นเวลา
        ไทยอยู่แล้ว" เพราะ connection ทุกเส้นถูกตั้ง SET TIME ZONE 'Asia/Bangkok' ไว้ที่ session ผ่าน
        core/apps.py's RLS signal (จุดเดียวกับที่ template อื่นๆ format วันที่ตรงๆ โดยไม่ต้องแปลงมาตลอด)
        — ใช้ค่าตรงๆ ได้เลย ห้ามเรียก timezone.localtime() ซ้ำ (จะ error เพราะเป็น naive datetime)
        """
        ms = self.created_at.strftime("%f")[:3]
        return f"แพลน_7-11_{self.created_at.strftime('%Y-%m-%d_%H:%M:%S')}:{ms}"


class PlanRunImport(models.Model):
    """ตารางกลาง (through) ระหว่าง PlanRun กับ PoImport — PK คู่ (plan_run_id, po_import_id) ในฐานข้อมูล
    จริง ใช้ CompositePrimaryKey ได้ตรงนี้เพราะไม่มีความจำเป็นต้องลงทะเบียนใน Django Admin เลย
    (ต่างจาก SkuMaster/LocationMapping ที่ต้องใช้ผ่าน Admin จึงติดข้อจำกัดเรื่อง composite PK)"""
    pk = models.CompositePrimaryKey("plan_run_id", "po_import_id")
    plan_run = models.ForeignKey(PlanRun, on_delete=models.CASCADE, db_column="plan_run_id")
    po_import = models.ForeignKey(PoImport, on_delete=models.CASCADE, db_column="po_import_id")

    class Meta:
        db_table = "plan_run_import"
        managed = False


class PlanSkuResult(models.Model):
    """
    ผลลัพธ์ต่อ SKU/คอลัมน์ของแผน (Phase 1.6 sub-phase 2) — "1 แถว = 1 SKU x 1 คอลัมน์" เก็บยอดสั่งจริง
    + ค่าที่ LibreOffice คำนวณจากสูตรจริงในเทมเพลต (ไม่ใช่สูตรจำลองใน Python) ใช้แสดงหน้าเว็บได้โดยตรง
    ไม่ต้องพึ่งไฟล์ Excel ที่อาจหายไปได้ (เช่นตอน Codespaces rebuild ที่เคยเจอปัญหานี้มาแล้วจริง)
    """
    plan_run = models.ForeignKey(PlanRun, on_delete=models.CASCADE, db_column="plan_run_id",
                                  related_name="sku_results")
    sheet_type = models.CharField(max_length=20)  # 'production' | 'logistic'
    group_name = models.CharField(max_length=50, blank=True, null=True)
    barcode = models.CharField(max_length=20)
    name_th = models.TextField(blank=True, null=True)
    name_en = models.TextField(blank=True, null=True)
    price = models.DecimalField(max_digits=10, decimal_places=2, blank=True, null=True)
    pack_size = models.IntegerField(blank=True, null=True)
    column_label = models.CharField(max_length=100)
    qty = models.DecimalField(max_digits=10, decimal_places=2, blank=True, null=True)
    pack_text = models.CharField(max_length=50, blank=True, null=True)
    grand_total = models.DecimalField(max_digits=12, decimal_places=2, blank=True, null=True)
    buffer_qty = models.DecimalField(max_digits=10, decimal_places=2, blank=True, null=True)
    return_qty = models.DecimalField(max_digits=10, decimal_places=2, blank=True, null=True)
    basket_total = models.IntegerField(blank=True, null=True)

    class Meta:
        db_table = "plan_sku_result"
        managed = False
        verbose_name = "Plan SKU Result"
        verbose_name_plural = "Plan SKU Results"

    def __str__(self):
        return f"{self.barcode} — {self.column_label}"


class PlanRunLogisticFile(models.Model):
    plan_run = models.ForeignKey(PlanRun, on_delete=models.CASCADE, db_column="plan_run_id",
                                  related_name="logistic_files")
    group_name = models.CharField(max_length=50)
    status = models.CharField(max_length=20)
    file_path = models.TextField(blank=True, null=True)
    error_message = models.TextField(blank=True, null=True)
    template_version = models.ForeignKey(
        TemplateVersion, on_delete=models.SET_NULL, db_column="template_version_id",
        blank=True, null=True, related_name="+",
    )

    class Meta:
        db_table = "plan_run_logistic_file"
        managed = False
        verbose_name = "Logistic File"
        verbose_name_plural = "Logistic Files"

    def __str__(self):
        return f"{self.group_name} ({self.status})"
