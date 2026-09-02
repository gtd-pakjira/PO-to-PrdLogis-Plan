#!/bin/bash
set -e

# ติดตั้ง LibreOffice (ใช้คำนวณสูตรใน .xlsx จริง — ดู customers/cpall/logic/excel_calc.py)
# ติดตั้งเฉพาะ libreoffice-calc (ไม่เอา Writer/Impress/Draw) เพื่อลดขนาด/เวลาติดตั้ง
if ! command -v soffice > /dev/null 2>&1; then
    echo "[postCreate] ติดตั้ง LibreOffice Calc (ใช้คำนวณสูตร Excel)..."
    sudo apt-get update -qq
    sudo apt-get install -y -qq --no-install-recommends libreoffice-calc
fi

# สร้าง .env อัตโนมัติถ้ายังไม่มี (ครั้งแรกที่เปิด Codespaces/devcontainer เท่านั้น) — generate
# SECRET_KEY ใหม่ทุกครั้งที่สร้างไฟล์นี้ ไม่ใช้ค่า placeholder ใน .env.example ตรงๆ เด็ดขาด
if [ ! -f .env ]; then
    python3 -c "
import secrets, string
chars = string.ascii_letters + string.digits + '!@#\$%^&*(-_=+)'
key = ''.join(secrets.choice(chars) for _ in range(50))
with open('.env.example') as f:
    content = f.read()
content = content.replace('django-insecure-เปลี่ยนค่านี้ก่อนใช้งานจริงเสมอ', key)
with open('.env', 'w') as f:
    f.write(content)
print('[postCreate] สร้าง .env ใหม่พร้อม SECRET_KEY สุ่มแล้ว')
"
fi

pip install -r requirements.txt --break-system-packages
python -c "from core.db import run_schema; run_schema('sql/schema.sql')"
python manage.py sync_cpall_config
DB_USER=postgres DB_PASSWORD=postgres python manage.py migrate
python -c "from core.db import run_schema; run_schema('sql/schema.sql')"

echo "[postCreate] เสร็จแล้ว — เหลือแค่ DB_USER=postgres DB_PASSWORD=postgres python manage.py createsuperuser (ถ้ายังไม่เคยทำ)"
