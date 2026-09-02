from django.apps import AppConfig


class CpallConfig(AppConfig):
    default_auto_field = "django.db.models.BigAutoField"
    name = "customers.cpall"
    label = "cpall"  # กำหนดชัดเจน (ไม่ใช้ default ที่ derive จาก name) กัน admin URL/app_name เปลี่ยนไป
    verbose_name = "CP All (7-11)"
