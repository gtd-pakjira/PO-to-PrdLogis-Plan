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
    id = models.AutoField(primary_key=True)
    barcode = models.CharField(max_length=20, verbose_name="บาร์โค้ด")
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
        verbose_name = "รหัสสินค้า"
        verbose_name_plural = "รหัสสินค้า (Product Code)"
        constraints = [
            models.UniqueConstraint(fields=["customer", "barcode"], name="sku_master_customer_barcode_key"),
        ]

    def __str__(self):
        return f"{self.barcode} — {self.name_th}"


class LocationMapping(models.Model):
    id = models.AutoField(primary_key=True)
    fc_code = models.CharField(max_length=10, verbose_name="รหัสสถานที่ (FC code)")
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
        constraints = [
            models.UniqueConstraint(fields=["customer", "fc_code"], name="location_mapping_customer_fc_code_key"),
        ]

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
    # ลำดับ+ชื่อคอลัมน์ทั้งหมดในไฟล์ต้นฉบับ (list) — ไฟล์จริงมีคอลัมน์ชื่อซ้ำกันได้ (เช่น "Discount
    # Percentage 1" ปรากฏ 2 รอบ) จับคู่กับ PoLine.all_values ด้วยตำแหน่ง ไม่ใช่ชื่อ ตอนสร้างไฟล์ใหม่
    column_order = models.JSONField(blank=True, null=True)

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
    # ค่าทุกคอลัมน์ของแถวนี้ตามลำดับเดียวกับ PoImport.column_order — เก็บไว้สร้างไฟล์ใหม่ให้ข้อมูล
    # ครบเหมือนต้นฉบับได้ แม้จะไม่ได้ใช้คอลัมน์เหล่านั้นในการคำนวณของระบบเลยก็ตาม
    all_values = models.JSONField(blank=True, null=True)

    class Meta:
        db_table = "po_line"
        managed = False
        verbose_name = "PO Line"
        verbose_name_plural = "PO Lines"

    def __str__(self):
        return f"{self.po_number} — {self.barcode}"


class LogisticGroup(models.Model):
    """
    กลุ่มพื้นที่ของ Logistic Plan (เดิม hardcode ไว้เป็น GROUP_TEMPLATES dict ในโค้ด — ย้ายมาเก็บเป็น
    ข้อมูลแทน เพิ่ม/แก้/ปิดใช้งานกลุ่มได้ผ่านหน้าเว็บโดยตรง ไม่ต้องแก้โค้ด+deploy ใหม่)
    """
    customer = models.ForeignKey(Customer, on_delete=models.CASCADE, db_column="customer_id")
    group_name = models.CharField(max_length=50, verbose_name="ชื่อกลุ่มพื้นที่")
    template_key = models.CharField(max_length=50, verbose_name="Template key (ต้องขึ้นต้นด้วย logistic_)")
    sheet_name = models.CharField(max_length=100, verbose_name="ชื่อ Sheet ในไฟล์เทมเพลต")
    display_order = models.IntegerField(default=0, verbose_name="ลำดับแสดงผล")
    is_active = models.BooleanField(default=True, verbose_name="เปิดใช้งาน")
    created_at = models.DateTimeField(auto_now_add=True)

    class Meta:
        db_table = "logistic_group"
        managed = False
        verbose_name = "กลุ่มพื้นที่ (Logistic)"
        verbose_name_plural = "กลุ่มพื้นที่ (Logistic)"
        ordering = ["display_order", "group_name"]
        constraints = [
            models.UniqueConstraint(fields=["customer", "group_name"], name="logistic_group_customer_name_key"),
            models.UniqueConstraint(fields=["customer", "template_key"], name="logistic_group_customer_key_key"),
        ]

    def __str__(self):
        return self.group_name

    def clean(self):
        from django.core.exceptions import ValidationError
        if self.template_key and not self.template_key.startswith("logistic_"):
            raise ValidationError({"template_key": "ต้องขึ้นต้นด้วย 'logistic_' เสมอ (เช่น 'logistic_ระยอง')"})


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
    original_filename = models.TextField(blank=True, null=True)
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

    def get_display_name(self, prefix="แพลน"):
        """
        ชื่อไฟล์ดาวน์โหลดตามฟอร์แมตที่ตกลงกันไว้: {prefix}_7-11_YYYY-MM-DD_HH-MM-SS-FF3
        prefix เปลี่ยนตามประเภทไฟล์ — "แพลน" สำหรับ Production Plan, ชื่อกลุ่ม (เช่น "บางบัวทอง")
        สำหรับ Logistic Plan — คำนวณจาก created_at ทุกครั้ง ไม่ได้เก็บเป็นคอลัมน์แยก

        ใช้ "-" คั่นเวลา (เดิมใช้ ":" ตามฟอร์แมต Postgres แต่ ":" ใช้ในชื่อไฟล์บน Windows ไม่ได้จริง
        แม้เบราว์เซอร์ส่วนใหญ่จะแปลงให้อัตโนมัติ แต่ user เจอปัญหาจริงเลยเปลี่ยนมาใช้ "-" ให้ชัวร์ 100%)

        หมายเหตุ: created_at ที่ได้จาก ORM เป็น naive datetime (ไม่มี tzinfo) แต่ "ค่าจริงเป็นเวลา
        ไทยอยู่แล้ว" เพราะ connection ทุกเส้นถูกตั้ง SET TIME ZONE 'Asia/Bangkok' ไว้ที่ session ผ่าน
        core/apps.py's RLS signal (จุดเดียวกับที่ template อื่นๆ format วันที่ตรงๆ โดยไม่ต้องแปลงมาตลอด)
        — ใช้ค่าตรงๆ ได้เลย ห้ามเรียก timezone.localtime() ซ้ำ (จะ error เพราะเป็น naive datetime)
        """
        ms = self.created_at.strftime("%f")[:3]
        return f"{prefix}_7-11_{self.created_at.strftime('%Y-%m-%d_%H-%M-%S')}-{ms}"

    def get_short_label(self):
        """ป้ายกำกับสั้นๆ ใช้แสดงในหน้าเว็บ (list/title): 'แผน #1 2026-09-02_11:54' (ไม่มีวินาที/ms
        ต่างจาก get_display_name() ที่ใช้ตั้งชื่อไฟล์ ต้องละเอียดกว่าเพื่อกันชื่อไฟล์ชนกัน)"""
        return f"แผน #{self.id} {self.created_at.strftime('%Y-%m-%d_%H:%M')}"


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
        verbose_name = "ผลคำนวณแผน (ต่อรหัสสินค้า)"
        verbose_name_plural = "ผลคำนวณแผน (ต่อรหัสสินค้า)"

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
