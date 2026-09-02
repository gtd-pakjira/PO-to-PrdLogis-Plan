"""
core/db.py — จัดการ connection ไปยัง PostgreSQL (ของกลาง ใช้ร่วมกันทุกลูกค้า)

มี 2 ฟังก์ชันเชื่อมต่อ แยกจุดประสงค์กันชัดเจน:

  get_connection(customer_id=...)  — connection ปกติสำหรับใช้งานจริงในแอป เชื่อมต่อด้วย role
      "app_role" (ไม่ใช่ superuser) และตั้งค่า session variable "app.current_customer_id" ทันที
      ให้ Row-Level Security (RLS) ของ Postgres บังคับกรองข้อมูลเฉพาะลูกค้านั้นอัตโนมัติ

  get_admin_connection()  — connection สำหรับรัน schema.sql/migration เท่านั้น เชื่อมต่อด้วย
      superuser "postgres" เพราะต้องมีสิทธิ์ทำ DDL (CREATE TABLE, ALTER TABLE, CREATE ROLE ฯลฯ)

*** สำคัญมาก ***
ห้ามใช้ get_admin_connection() ทำ query ข้อมูลทั่วไปเด็ดขาด เพราะ superuser "ข้าม" RLS เสมอ
(Postgres ไม่บังคับ RLS กับ superuser ไม่ว่าจะตั้ง FORCE ROW LEVEL SECURITY หรือไม่ก็ตาม) —
ถ้าใช้ get_admin_connection() สำหรับ query ทั่วไป จะเห็นข้อมูลข้ามลูกค้าได้หมด ทำลาย RLS ที่ตั้งไว้

ตั้งค่าผ่าน environment variables:
    DB_HOST=localhost, DB_PORT=5432, DB_NAME=gtd_poc
    DB_USER=app_role, DB_PASSWORD=app_password           (สำหรับ get_connection — ปกติ)
    DB_ADMIN_USER=postgres, DB_ADMIN_PASSWORD=postgres    (สำหรับ get_admin_connection — schema เท่านั้น)
"""
import os

import psycopg2
import psycopg2.extras


def _connect(user: str, password: str):
    return psycopg2.connect(
        host=os.environ.get("DB_HOST", "localhost"),
        port=os.environ.get("DB_PORT", "5432"),
        dbname=os.environ.get("DB_NAME", "gtd_poc"),
        user=user,
        password=password,
    )


def get_connection(customer_id: int | None = None):
    """
    เปิด connection ใหม่สำหรับใช้งานจริง (role ไม่ใช่ superuser — โดน RLS บังคับจริง)
    customer_id: ถ้าระบุมา จะตั้ง session variable ให้ RLS กรองเฉพาะข้อมูลลูกค้านั้น
    ถ้าไม่ระบุ (None) — RLS จะกันไม่ให้เห็นข้อมูลอะไรเลยในตารางที่มี RLS (ปลอดภัยไว้ก่อนโดย default)
    """
    conn = _connect(
        os.environ.get("DB_USER", "app_role"),
        os.environ.get("DB_PASSWORD", "app_password"),
    )
    with conn.cursor() as cur:
        cur.execute("SET TIME ZONE 'Asia/Bangkok'")
        if customer_id is not None:
            cur.execute("SET app.current_customer_id = %s", (str(customer_id),))
    conn.commit()
    return conn


def get_admin_connection():
    """เปิด connection แบบ superuser — ใช้ได้เฉพาะตอนรัน schema.sql/migration เท่านั้น ห้ามใช้ query ข้อมูลทั่วไป"""
    conn = _connect(
        os.environ.get("DB_ADMIN_USER", "postgres"),
        os.environ.get("DB_ADMIN_PASSWORD", "postgres"),
    )
    with conn.cursor() as cur:
        cur.execute("SET TIME ZONE 'Asia/Bangkok'")
    conn.commit()
    return conn


def run_schema(schema_path: str):
    """รันไฟล์ sql/schema.sql เพื่อสร้าง/migrate ตารางทั้งหมด (idempotent, รันซ้ำได้) — ใช้ admin connection เสมอ"""
    with open(schema_path, "r", encoding="utf-8") as f:
        sql = f.read()
    conn = get_admin_connection()
    try:
        with conn.cursor() as cur:
            cur.execute(sql)
        conn.commit()
    finally:
        conn.close()
