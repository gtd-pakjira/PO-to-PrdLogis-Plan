"""
cpall/logic/db.py — ตัวกลางระหว่าง business logic ของ cpall กับ core.db

จุดสำคัญ: ทุกไฟล์ใน cpall/logic/*.py เดิม import "get_connection" จากที่นี่ (แทนที่จะ import
core.db ตรงๆ) — ฟังก์ชันนี้ resolve customer_id ของ "cpall" ให้อัตโนมัติแล้วส่งต่อให้
core.db.get_connection(customer_id=...) ทำให้ RLS ทำงานถูกต้องโดยที่ไม่ต้องแก้โค้ดเดิมนับสิบไฟล์
เลยสักบรรทัด (แค่เปลี่ยน import 1 บรรทัดต่อไฟล์ จาก "from src.db" เป็น "from customers.cpall.logic.db")
"""
from core.db import get_connection as _core_get_connection

_CPALL_CUSTOMER_ID = None  # cache ไว้ ไม่ต้อง query ซ้ำทุกครั้ง


def _resolve_cpall_customer_id() -> int:
    global _CPALL_CUSTOMER_ID
    if _CPALL_CUSTOMER_ID is None:
        # การ query ครั้งนี้ใช้ get_connection(customer_id=None) ซึ่งปกติ RLS จะบล็อกไม่ให้เห็น
        # อะไรเลยในตารางที่มี RLS — แต่ตาราง "customer" ไม่มี RLS (ดู schema.sql) จึง query ได้ปกติ
        conn = _core_get_connection()
        try:
            with conn.cursor() as cur:
                cur.execute("SELECT id FROM customer WHERE code = 'cpall'")
                row = cur.fetchone()
        finally:
            conn.close()
        if row is None:
            raise RuntimeError(
                "ไม่พบลูกค้า 'cpall' ในตาราง customer — ต้องรัน sql/schema.sql ก่อน (มี seed data อยู่แล้ว)"
            )
        _CPALL_CUSTOMER_ID = row[0]
    return _CPALL_CUSTOMER_ID


def get_connection():
    """เหมือน core.db.get_connection() ทุกอย่าง แต่ผูก customer_id ของ cpall ให้อัตโนมัติเสมอ"""
    return _core_get_connection(customer_id=_resolve_cpall_customer_id())


def get_cpall_customer_id() -> int:
    """เผยค่า customer_id ของ cpall ออกมาให้ใช้ตรงๆ ได้ (เช่น ตอน INSERT ที่ต้องใส่ค่าใน VALUES เอง)"""
    return _resolve_cpall_customer_id()
