# แพลนผลิต 7-11 — Django, โครงสร้างแบบ "1 ลูกค้า = 1 โมดูล"

> ตัดสินใจใช้ Django ต่อจาก POC เปรียบเทียบ Flask/Django (2026-08-29) — เหตุผลหลัก: ต้องรองรับ
> หลายลูกค้าในอนาคต + ต้องมีระบบสิทธิ์ผู้ใช้งาน ซึ่ง Django มี Admin + Auth/Permission ในตัวพร้อมใช้
>
> รื้อโครงสร้างรอบใหญ่ (2026-08-30) เพื่อรองรับหลายลูกค้าจริงจัง: แยกโค้ดเป็น "1 ลูกค้า = 1 Django app"
> (ตอนนี้มีแค่ `cpall` = CP All/7-11) ผ่าน Portal ที่เลือกลูกค้าได้ + Postgres Row-Level Security (RLS)
> กันข้อมูลรั่วข้ามลูกค้าที่ระดับฐานข้อมูลเอง (ไม่ใช่แค่พึ่งโค้ดฝั่งแอปอย่างเดียว)

## สถาปัตยกรรม

```
core/            ของกลาง ใช้ร่วมกันทุกลูกค้า
  db.py            connection helper — แยก 2 แบบชัดเจน:
                     get_connection(customer_id=..) = role จำกัดสิทธิ์ ใช้งานจริง โดน RLS บังคับ
                     get_admin_connection()          = superuser ใช้แค่ตอนรัน schema.sql เท่านั้น
  models.py        Customer (ตารางกลาง ไม่มี RLS เพราะไม่ใช่ข้อมูลอ่อนไหว)

portal/          หน้าแรกสุด (/) — ภาพรวมทุกลูกค้า + ปุ่มเข้าแต่ละเจ้า

cpall/           โมดูลของลูกค้า CP All (7-11) ทั้งหมด — ตัวอย่าง/ต้นแบบสำหรับลูกค้าเจ้าถัดไป
  logic/           business logic ล้วนๆ (ไม่ใช่ Django-specific) — ใช้ raw SQL ผ่าน psycopg2
    db.py            resolve customer_id ของ cpall อัตโนมัติ แล้วส่งต่อให้ core.db
    po_parser.py, grouping.py, excel_export.py, logistic_plan_export.py, plan_runner.py,
    plan_view_data.py, template_manager.py, location_mapping_manager.py, config_loader.py, date_utils.py
  models.py        Django ORM model ของ SkuMaster/LocationMapping (สำหรับ Django Admin เท่านั้น)
  views.py, urls.py, admin.py
  templates/cpall/
  config/          sku_master.yaml, location_mapping.yaml (เฉพาะ cpall)
  excel_templates/ ไฟล์ Template Excel ทั้ง 5 ไฟล์ (เฉพาะ cpall)
  data/            po_uploads/, output/ (เฉพาะ cpall)

sql/schema.sql   ตาราง Postgres ทั้งหมด (core + cpall) + RLS policy + app_role
webproject/      Django project settings/urls
```

## Row-Level Security (RLS) — ทำไมต้องมี 2 role

**`app_role`** (ไม่ใช่ superuser) คือ role ที่แอปใช้เชื่อมต่อจริงตอนรันงาน — Postgres จะบังคับ RLS
กับ role นี้จริง (กรองให้เห็นแค่ข้อมูลของ `customer_id` ที่ session ตั้งไว้เท่านั้น แม้โค้ดจะลืม
`WHERE customer_id = ...` เอง ก็ยังปลอดภัย)

**`postgres`** (superuser) ใช้ได้แค่ตอนรัน `sql/schema.sql`/`migrate` เท่านั้น — **ห้ามใช้ query ข้อมูล
ทั่วไปเด็ดขาด** เพราะ Postgres ไม่บังคับ RLS กับ superuser ไม่ว่าจะตั้ง `FORCE ROW LEVEL SECURITY`
หรือไม่ก็ตาม (พิสูจน์เจอบั๊กนี้จริงระหว่างพัฒนา — ตอนแรกลืมจุดนี้ RLS เลยไม่มีผลอะไรเลย)

**ข้อจำกัดของ Django Admin ที่ควรรู้:** `SkuMaster`/`LocationMapping` ใน DB จริงมี PK แบบ
`(customer_id, barcode)` composite เพื่อรองรับหลายลูกค้าใช้บาร์โค้ดซ้ำกันได้ — แต่ Django Admin
**ลงทะเบียน model ที่มี composite primary key ไม่ได้** (ข้อจำกัดของ Django) เลยให้ `barcode`/`fc_code`
เป็น Django PK เดี่ยวไปก่อน ใช้ได้ปกติเพราะตอนนี้มีลูกค้าเดียว — ถ้าวันหน้ามีลูกค้าที่ 2 ที่บาร์โค้ดซ้ำกัน
จริงๆ ต้องเปลี่ยนไปใช้ surrogate key (เพิ่มคอลัมน์ `id` ในตาราง DB จริง)

## สถานะปัจจุบัน

**เสร็จแล้ว:**
- Portal (`/`) — ภาพรวมลูกค้าทั้งหมด, cpall Dashboard (`/cpall/`) — ครบทุกฟีเจอร์เดิม
- นำเข้า PO (เก็บชื่อไฟล์เดิมไว้ ไม่เปลี่ยนชื่อทิ้ง) + ตัดแถวข้อมูลซ้ำเป๊ะภายในไฟล์เดียวกันอัตโนมัติ
- เลือก location mapping ผ่านเว็บเมื่อเจอรหัสสถานที่ใหม่ (บันทึกทั้ง DB + YAML)
- กรอกยอดเผื่อผ่านเว็บก่อนสร้างแผน (ดึงค่าเริ่มต้นจากเทมเพลต)
- สร้างแผน (Production Plan + Logistic Plan 4 กลุ่ม) ตรงกับไฟล์จริงของ Admin 100%
- ดูตารางตัวเลขในเว็บ + ดาวน์โหลดไฟล์ Excel จริง
- จัดการ Template (ดาวน์โหลด/อัปโหลดแทนที่ พร้อม validate โครงสร้าง + backup อัตโนมัติ)
- ลบ PO / ลบแผน (มีกันลบ PO ที่ถูกใช้สร้างแผนไปแล้ว)
- Django Admin จัดการ SKU/Location/Customer

**ยังไม่ได้ทำ (คงค้างจากรอบก่อนๆ หรือเพิ่งคุยกันไว้):**
- [ ] จัดรถ/box-fill — รอ user อัปเดตข้อมูล
- [ ] ระบบสิทธิ์ผู้ใช้งาน (role/permission) — เก็บไว้ทำทีหลังตามที่ตกลงกัน (โครงสร้างโมดูล+RLS ที่ทำ
      รอบนี้รองรับได้เลยตอนถึงเวลาทำจริง)
- [ ] เปลี่ยนจาก "เก็บเป็นไฟล์ Excel" เป็น "เก็บเป็นข้อมูล แล้วค่อย generate ไฟล์ตอนกดดาวน์โหลด" —
      ไอเดียดีที่คุยกันไว้ แต่เป็นงานใหญ่แยกต่างหาก ยังไม่เริ่ม
- [ ] ตรวจจับ "อัปโหลดไฟล์ PO เดิมซ้ำคนละรอบ" (ต่างจาก duplicate ภายในไฟล์เดียวที่ทำไปแล้ว) — ยังไม่ได้
      ออกแบบว่า "ซ้ำ" ควรนิยามยังไง (เลข PO ซ้ำ? เนื้อหาไฟล์เหมือนกันเป๊ะ?)
- [ ] `main.py` (CLI) ไม่มีในเวอร์ชัน Django แล้ว (ใช้ผ่านเว็บทั้งหมด) — ถ้าต้องการ CLI กลับมาบอกได้
- [ ] UI redesign (sidebar, Stripe-style table/badge) ที่คุยกันไว้ก่อนเริ่มรื้อโครงสร้าง — ยังไม่ได้ทำ
      (ทำโครงสร้างเสร็จก่อนตามลำดับความสำคัญที่ตกลงกัน)

## วิธีตั้งค่าบน GitHub Codespaces

1. Push โฟลเดอร์นี้ (รวม `.devcontainer/`) ขึ้น GitHub repo — **ตรวจสอบว่ามี `.gitignore` ติดไปด้วย
   ก่อน push ครั้งแรกเสมอ** (กันไฟล์ PO/แผนจริงกับ `.env` หลุดเข้า git โดยไม่ตั้งใจ)
2. "Code" → "Codespaces" → "Create codespace on main" — รอสักครู่ ระบบจะรัน `.devcontainer/postCreate.sh`
   ให้อัตโนมัติ: **สร้าง `.env` พร้อม `SECRET_KEY` สุ่มใหม่ให้เอง** (ถ้ายังไม่มีไฟล์ `.env`), ติดตั้ง
   Python library, รัน `schema.sql` (สร้างตาราง + `app_role` + RLS), sync config จาก YAML, รัน Django
   `migrate`, รัน `schema.sql` อีกรอบ (grant สิทธิ์ `app_role` ให้เห็นตาราง Django ที่เพิ่งสร้าง)
3. สร้าง superuser สำหรับเข้า `/admin/` (ครั้งแรกครั้งเดียว):
   ```bash
   DB_USER=postgres DB_PASSWORD=postgres python manage.py createsuperuser
   ```
   (ต้องสลับไปใช้ `postgres` ชั่วคราว เพราะ `createsuperuser` เขียนตาราง auth ของ Django)
4. รันเว็บ (ใช้ `app_role` ตามค่า default ใน `.devcontainer` อยู่แล้ว ไม่ต้องตั้งอะไรเพิ่ม):
   ```bash
   python manage.py runserver 0.0.0.0:8000
   ```
5. เปิด (ดูแท็บ "PORTS" หาพอร์ต 8000): `/` = Portal, `/cpall/` = ระบบ 7-11, `/admin/` = Django Admin

### รันนอก Codespaces (เช่น บนโน้ตบุ๊กเครื่องใหม่ที่ไม่ได้ใช้ Codespaces)
`.devcontainer/postCreate.sh` สร้าง `.env` ให้อัตโนมัติแค่ตอนเปิดผ่าน Codespaces/devcontainer เท่านั้น
— ถ้ารันนอก devcontainer ต้องสร้างเองครั้งแรก:
```bash
cp .env.example .env
python -c "from django.core.management.utils import get_random_secret_key; print(get_random_secret_key())"
# เอาผลลัพธ์ไปแทนค่า SECRET_KEY ใน .env
```

### pgAdmin (ดูข้อมูลแบบคลิกดู)
แท็บ "PORTS" → พอร์ต **5050** → login `admin@example.com` / `admin12345` → Register Server →
Host=`db`, Port=`5432`, Username=`postgres`, Password=`postgres`, Database=`gtd_poc`

(ถ้ายัง login ไม่ได้หลังอัปเดตนี้ — ลอง "Rebuild Container" เพื่อให้ pgAdmin container สร้างใหม่รับค่า
password ใหม่ ถ้ายังไม่ได้อีก แจ้งกลับมาเพื่อ debug ต่อ ยังไม่เคยทดสอบกับ Docker จริง มีแค่การวิเคราะห์
จาก config)

### ⚠️ ต้องทำเองก่อน push ครั้งต่อไป (ผมทำให้ไม่ได้ เพราะเข้าไม่ถึง git repo จริงของคุณ)

Repo นี้เคย commit ไฟล์ที่ไม่ควรมีเข้า git ตั้งแต่ commit แรก (ก่อนมี `.gitignore`) — ต้องล้าง git
history ทิ้งแล้วเริ่มใหม่ก่อน push รอบถัดไป (repo เป็น Private คนเดียวใช้ ความเสี่ยงต่ำ แต่ควรทำให้ถูกต้อง):

```bash
# 1. เอาไฟล์ชุดนี้ (ที่มี .gitignore ใหม่แล้ว) ไปวางทับของเดิมก่อน
rm -rf .git
git init
git add .
git commit -m "Reset history — add .gitignore, django-environ, Ruff (Phase 0)"
git remote add origin <URL repo เดิมของคุณ>
git branch -M main
git push --force origin main
```

**ทำครั้งเดียวจบ** หลังจากนี้ commit ปกติได้เลยไม่ต้อง `--force` อีก

## เครื่องมือช่วยตรวจโค้ด
```bash
pip install -r requirements-dev.txt --break-system-packages
ruff check .          # ตรวจสอบ
ruff check . --fix    # แก้ที่แก้อัตโนมัติได้
```

## ข้อจำกัดที่ควรรู้ก่อนใช้งานจริง

- **ยอดเผื่อยังไม่มีสูตรคำนวณจริง** — หน้าเว็บให้กรอกเองได้แล้ว แต่ค่าเริ่มต้นยังดึงจากไฟล์เทมเพลต
  `cpall/excel_templates/logistic_plan_รอบเช้าต่างจังหวัด.xlsx` (Admin เคยกรอกไว้) ต้องถาม Admin จริง
  ว่าคำนวณยังไง หรือต้องกรอกมือเสมอ
- **ลำดับ PO1/PO2/PO3...** เดาจากเรียงเลข PO น้อยไปมาก — ยังไม่ได้ยืนยันกับ Admin
- **ไฟล์เทมเพลต "รอบเช้าต่างจังหวัด" มีแค่ 18 SKU** (ขาดพุทราจีน) ต่างจาก 3 กลุ่มอื่นที่มี 19
- **จำนวนคอลัมน์ PO ต่อจุดส่งในเทมเพลต Logistic Plan ไม่คงที่** — เกินแล้วต้องไปเพิ่มคอลัมน์ในเทมเพลตเอง
  (ใช้หน้า "จัดการ Template" ในเว็บได้)

## คำถามที่ต้องเตรียมถาม Admin (Note ถึงตัวเอง)

- **ยอดเผื่อ**: คำนวณยังไง? มีสูตรจริงไหม หรือ Admin ตัดสินใจเองล้วนๆ ทุกครั้ง? ทำไมมีแค่ไฟล์
  "รอบเช้าต่างจังหวัด" ไฟล์เดียวที่มีคอลัมน์นี้?
- **ลำดับ PO1/PO2/PO3**: Admin เรียงเข้าคอลัมน์ตามอะไรจริงๆ?
- **พุทราจีน**: ทำไมเทมเพลต "รอบเช้าต่างจังหวัด" ไม่มีแถวนี้? ปกติจุดกลุ่มนี้ไม่เคยสั่งพุทราจีนเลยจริงไหม?
- **จำนวนคอลัมน์ PO ต่อจุดส่ง**: มีขั้นตอน/กฎเกณฑ์ตายตัวไหมตอนเพิ่มคอลัมน์ใหม่?
- **การจัดรถ**: ความจุตะกร้า/ประเภทรถมีกี่แบบ? สูตร box-fill (F / >H / ≤H) คำนวณจากอะไรบ้าง?
  การมอบหมายรถต่อกลุ่มพื้นที่คงที่เสมอ หรือเปลี่ยนได้ตามยอด/วัน?
- **จุดส่ง/SKU ใหม่ในอนาคต**: ต้องทำยังไง ต้องแก้ที่ไหนบ้าง?
