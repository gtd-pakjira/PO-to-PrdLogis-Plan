"""
forms.py — Django Forms สำหรับ validate ข้อมูลจากฟอร์มในเว็บ (Phase 1)

เดิมทุกฟอร์มอ่านค่าจาก request.POST/request.FILES ตรงๆ ไม่ผ่านการ validate จริงจังเลย (เช็คแค่บางจุด
เช่น ลองแปลงวันที่แล้ว catch ValueError) — ย้ายมาใช้ Django Forms เพื่อให้ validate แบบมาตรฐาน มี error
message ต่อฟิลด์ชัดเจน และ render ด้วย django-crispy-forms (ธีม Tailwind) ให้หน้าตาสม่ำเสมอกับที่มีอยู่แล้ว

ฟอร์มที่มีจำนวนฟิลด์ไม่แน่นอน (กรอกยอดเผื่อ, เลือก location mapping) ยังใช้วิธีอ่านจาก request.POST
ตรงๆ ต่อไปก่อน เพราะ field สร้างจากข้อมูล runtime (จำนวน SKU/จำนวนรหัสที่ไม่รู้จักไม่แน่นอน) — ทำเป็น
Django Form แบบ dynamic fields ได้เหมือนกัน แต่เพิ่มความซับซ้อนไม่คุ้มกับฟอร์มที่ไม่มีเงื่อนไข validate
พิเศษอะไรเลย (แค่ต้องเป็นตัวเลข) — เก็บไว้เป็นตัวเลือกทำเพิ่มทีหลังถ้าจำเป็น
"""
import os

from django import forms


class ImportPOForm(forms.Form):
    po_file = forms.FileField(
        label="ไฟล์ PO (.xlsx)",
        widget=forms.ClearableFileInput(attrs={"accept": ".xlsx"}),
    )
    production_date = forms.DateField(
        label="วันที่ผลิต",
        widget=forms.DateInput(attrs={"type": "date"}),
    )
    po_date = forms.DateField(
        label="วันที่ PO",
        widget=forms.DateInput(attrs={"type": "date"}),
    )

    def clean_po_file(self):
        po_file = self.cleaned_data["po_file"]
        ext = os.path.splitext(po_file.name)[1].lower()
        if ext != ".xlsx":
            raise forms.ValidationError("ต้องเป็นไฟล์ .xlsx เท่านั้น")
        return po_file


class TemplateUploadForm(forms.Form):
    template_file = forms.FileField(
        label="ไฟล์ Template ใหม่ (.xlsx)",
        widget=forms.ClearableFileInput(attrs={"accept": ".xlsx"}),
    )

    def clean_template_file(self):
        f = self.cleaned_data["template_file"]
        ext = os.path.splitext(f.name)[1].lower()
        if ext != ".xlsx":
            raise forms.ValidationError("ต้องเป็นไฟล์ .xlsx เท่านั้น")
        return f
