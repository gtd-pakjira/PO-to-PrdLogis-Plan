from django.apps import AppConfig
from django.db.backends.signals import connection_created


def _set_rls_session_vars(sender, connection, **kwargs):
    """
    ตั้งค่า session variable ให้ RLS ของ Postgres ทำงานถูกต้องบน connection ของ Django ORM เอง
    (Django Admin ใช้ connection นี้ตอนเข้าไปดู/แก้ SkuMaster, LocationMapping ที่มี RLS ป้องกันอยู่)

    *** ข้อจำกัดชั่วคราว *** ตอนนี้ระบบมีลูกค้าเดียว (cpall) เลย hardcode ให้ทุก connection ของ
    Django ORM มองเห็นแค่ cpall เสมอ — พอทำระบบสิทธิ์ผู้ใช้ในอนาคต ต้องเปลี่ยนจุดนี้ให้ดึง
    customer_id จาก request/session ของผู้ใช้แต่ละคนแทน ไม่ใช่ hardcode แบบนี้อีกต่อไป
    """
    if connection.vendor != "postgresql":
        return
    with connection.cursor() as cursor:
        cursor.execute("SET TIME ZONE 'Asia/Bangkok'")
        cursor.execute("SELECT id FROM customer WHERE code = 'cpall'")
        row = cursor.fetchone()
        if row:
            cursor.execute("SET app.current_customer_id = %s", (str(row[0]),))


class CoreConfig(AppConfig):
    default_auto_field = "django.db.models.BigAutoField"
    name = "core"

    def ready(self):
        connection_created.connect(_set_rls_session_vars)
