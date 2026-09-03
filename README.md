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

customers/cpall/ โมดูลของลูกค้า CP All (7-11) ทั้งหมด — ตัวอย่าง/ต้นแบบสำหรับลูกค้าเจ้าถัดไป
  logic/           business logic ล้วนๆ (ไม่ใช่ Django-specific) — ใช้ Django ORM เป็นหลัก (grouping.py
                     เป็นข้อยกเว้นที่ตั้งใจใช้ raw SQL ต่อ เพราะเป็น query รวม/JOIN ที่ซับซ้อน)
    db.py            resolve customer_id ของ cpall อัตโนมัติ แล้วส่งต่อให้ core.db
    po_parser.py, po_regenerator.py, po_view_data.py, grouping.py, excel_export.py,
    logistic_plan_export.py, plan_runner.py, plan_regenerator.py, plan_result_extractor.py,
    plan_view_data.py, template_manager.py, location_mapping_manager.py, config_loader.py, date_utils.py
  models.py        Django ORM model ทั้งหมด (SkuMaster, LocationMapping, LogisticGroup, PoImport,
                     PoLine, PlanRun, TemplateVersion, PlanSkuResult ฯลฯ) — ORM เต็มรูปแบบแล้ว ไม่ใช่
                     แค่สำหรับ Django Admin อีกต่อไป
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
- Portal (`/`) — ภาพรวมลูกค้าทั้งหมด, cpall Dashboard (`/cpall/`) — ครบทุกฟีเจอร์
- นำเข้า PO — เก็บทุกคอลัมน์ของไฟล์ต้นฉบับลง database ครบ (ไม่ใช่แค่ 12 คอลัมน์ที่ใช้คำนวณ) แล้ว
  **ลบไฟล์ต้นฉบับทิ้งทันที** (data-first) — ดู/ดาวน์โหลดย้อนหลังสร้างไฟล์ใหม่จาก database ตรงกับ
  ต้นฉบับทุกเซลล์ (ทดสอบเทียบแล้ว) + ตัดแถวข้อมูลซ้ำเป๊ะภายในไฟล์เดียวกันอัตโนมัติ
- เลือก location mapping ผ่านเว็บเมื่อเจอรหัสสถานที่ใหม่ (บันทึกทั้ง DB + YAML)
- กรอกยอดเผื่อผ่านเว็บก่อนสร้างแผน (ดึงค่าเริ่มต้นจากเทมเพลต)
- สร้างแผน (Production Plan + Logistic Plan ทุกกลุ่ม) ตรงกับไฟล์จริงของ Admin 100% — **ไฟล์ Excel ที่
  สร้างตอนแผนสำเร็จถูกลบทิ้งทันที** (data-first เต็มรูปแบบ) เว็บอ่านจาก database ล้วนๆ ดาวน์โหลด
  regenerate จากเทมเพลต+database สดทุกครั้ง (มีสูตรจริงในไฟล์ที่ได้เสมอ)
- โหลดทั้งแผนเป็น ZIP เดียว (Production + Logistic ทุกกลุ่ม)
- ดูตารางตัวเลขในเว็บ พร้อม tooltip เลข PO จริงเบื้องหลัง "PO1"/"PO2", แถวรวมยอดตะกร้าท้ายตาราง
- จัดการ Template (ดาวน์โหลด/อัปโหลดแทนที่ พร้อม validate โครงสร้าง + เก็บทุกเวอร์ชันไว้ถาวร กู้คืนได้)
- **กลุ่มพื้นที่ Logistic Plan ตั้งค่าได้ผ่าน Django Admin โดยตรง** (`/admin/cpall/logisticgroup/`)
  — เพิ่มกลุ่มที่ 5 ขึ้นไปได้เลยไม่ต้องแก้โค้ด/deploy ใหม่ (เดิม hardcode เป็น `GROUP_TEMPLATES` dict)
- ลบ PO / ลบแผน (กันลบ PO ที่ถูกใช้สร้างแผนไปแล้ว ต้องลบแผนก่อน — ไม่มีปุ่ม "บังคับลบ" แล้ว เพราะ PO
  เป็น data-first ไม่มีไฟล์ให้หายอีกต่อไป)
- HTMX ทั่วทั้งเว็บ — action ต่างๆ (ลบ/สร้างแผน/อัปโหลด) ไม่ reload ทั้งหน้า มี spinner ระหว่างรอ
- Django Admin จัดการ SKU/Location/Customer/กลุ่มพื้นที่ — ใช้ surrogate key (`id`) แทน composite PK
  เดิมแล้ว (sku_master/location_mapping)
- Kubi admin theme ปรับแต่งแล้ว (โลโก้, ลิงก์กลับ Portal, ปุ่มบันทึกเหลือชุดเดียว)

**ยังไม่ได้ทำ (ตั้งใจเก็บไว้ทีหลัง):**
- [ ] จัดรถ/box-fill — รอ user อัปเดตข้อมูล
- [ ] ระบบสิทธิ์ผู้ใช้งาน (role/permission) — โครงสร้างโมดูล+RLS รองรับได้เลยตอนถึงเวลาทำจริง
- [ ] ระบบ Log/Audit trail ฝั่งเว็บหลัก (cpall) — ยังไม่มีเลยสักจุด (ต้องรอระบบสิทธิ์ก่อน เพื่อให้รู้ว่า
      "ใคร" ทำ) — ที่มีอยู่ตอนนี้คือ History ของ Django Admin เอง (built-in ไม่ได้เขียนเพิ่ม) ซึ่งบันทึก
      เฉพาะการกระทำผ่าน `/admin/` เท่านั้น ไม่ครอบคลุมฝั่งเว็บหลักที่ User จริงใช้งาน
- [ ] ตรวจจับ "อัปโหลดไฟล์ PO เดิมซ้ำคนละรอบ" — ยังไม่ได้ออกแบบว่า "ซ้ำ" ควรนิยามยังไง
- [ ] `main.py` (CLI) ไม่มีในเวอร์ชัน Django แล้ว (ใช้ผ่านเว็บทั้งหมด)

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
  "รอบเช้าต่างจังหวัด" (Admin เคยกรอกไว้) ต้องถาม Admin จริงว่าคำนวณยังไง หรือต้องกรอกมือเสมอ
- **ลำดับ PO1/PO2/PO3...** เดาจากเรียงเลข PO น้อยไปมาก — ยังไม่ได้ยืนยันกับ Admin
- **ไฟล์เทมเพลต "รอบเช้าต่างจังหวัด" มีแค่ 18 SKU** (ขาดพุทราจีน) ต่างจากกลุ่มอื่นที่มี 19
- **จำนวนคอลัมน์ PO ต่อจุดส่งในเทมเพลต Logistic Plan ไม่คงที่** — เกินแล้วต้องไปเพิ่มคอลัมน์ในเทมเพลตเอง
  (ใช้หน้า "จัดการ Template" ในเว็บได้) ส่วนการ "เพิ่มกลุ่มพื้นที่ใหม่ทั้งกลุ่ม" ทำผ่าน Django Admin
  ได้แล้ว (`/admin/cpall/logisticgroup/`) ไม่ต้องแก้โค้ด — แต่ยังต้องเตรียมไฟล์เทมเพลตของกลุ่มนั้นเอง
  แล้วอัปโหลดที่หน้า "จัดการ Template" ให้ครบก่อนถึงจะสร้างแผนที่มีกลุ่มนั้นได้
- **ไฟล์ PO/Excel ที่สร้างระหว่างทำงานถูกลบทิ้งอัตโนมัติ** (data-first) — ถ้าต้องดูไฟล์จริงที่ Admin
  อัปโหลด/ระบบสร้างไว้ ต้องกดดูผ่านเว็บ (ระบบสร้างไฟล์ใหม่ให้สดๆ ตอนกดดาวน์โหลดเสมอ) ไม่มีไฟล์ค้างอยู่
  บนดิสก์ให้เปิดตรงๆ นอกเว็บอีกต่อไป — ยกเว้น PO/แผนเก่าที่สร้างก่อนอัปเดตนี้ ซึ่งยังมีไฟล์เดิมค้างอยู่

## คำถามที่ต้องเตรียมถาม Admin (Note ถึงตัวเอง)

- **ยอดเผื่อ**: คำนวณยังไง? มีสูตรจริงไหม หรือ Admin ตัดสินใจเองล้วนๆ ทุกครั้ง? ทำไมมีแค่ไฟล์
  "รอบเช้าต่างจังหวัด" ไฟล์เดียวที่มีคอลัมน์นี้?
- **ลำดับ PO1/PO2/PO3**: Admin เรียงเข้าคอลัมน์ตามอะไรจริงๆ?
- **พุทราจีน**: ทำไมเทมเพลต "รอบเช้าต่างจังหวัด" ไม่มีแถวนี้? ปกติจุดกลุ่มนี้ไม่เคยสั่งพุทราจีนเลยจริงไหม?
- **จำนวนคอลัมน์ PO ต่อจุดส่ง**: มีขั้นตอน/กฎเกณฑ์ตายตัวไหมตอนเพิ่มคอลัมน์ใหม่?
- **การจัดรถ**: ความจุตะกร้า/ประเภทรถมีกี่แบบ? สูตร box-fill (F / >H / ≤H) คำนวณจากอะไรบ้าง?
  การมอบหมายรถต่อกลุ่มพื้นที่คงที่เสมอ หรือเปลี่ยนได้ตามยอด/วัน?
- **จุดส่ง/SKU ใหม่ในอนาคต**: ต้องทำยังไง ต้องแก้ที่ไหนบ้าง?
