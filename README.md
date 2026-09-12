# PO-to-PrdLogis-Plan

ระบบจัดทำ **แพลนผลิต (Production Plan)** และ **แพลนกระจาย (Logistic Plan)** จากข้อมูล PO สำหรับงาน 7-11 / CP All

ระบบถูกออกแบบมาเพื่อ **ช่วย Workflow การทำงานที่มีอยู่จริง** ไม่ได้มีเป้าหมายให้ผู้ใช้งานต้องเปลี่ยนวิธีทำงานเพียงเพราะข้อจำกัดของระบบ

> **แนวคิดหลัก**
>
> ระบบควรปรับเข้าหา Workflow ของผู้ใช้งาน
> ไม่ใช่บังคับให้ผู้ใช้งานปรับ Workflow เข้าหาระบบ
>
> ระบบควรบังคับเฉพาะสิ่งที่จำเป็นต่อ Business Rule และ Data Integrity

---

## 1. เป้าหมายของระบบ

กระบวนการเดิมอาศัย PO และ Excel Template เป็นหลัก โดยผู้ใช้งานต้องจัดข้อมูล คำนวณ และเตรียมไฟล์สำหรับฝ่าย Production / Logistic ด้วยตนเอง

ระบบนี้เข้ามาช่วยในส่วนที่เหมาะกับการทำเป็นระบบ เช่น

* นำเข้า PO
* จัดการข้อมูล PO
* ตรวจสอบข้อมูลที่จำเป็น
* จัดการ Location / Sub-location
* คำนวณแพลน
* กำหนดและแก้ไขยอดเผื่อ
* สร้าง Production Plan
* สร้าง Logistic Plan
* แสดงผลบน Web
* สร้าง Excel จาก Template
* เก็บข้อมูลใน Database เพื่อให้สามารถสร้างไฟล์ใหม่ภายหลังได้

ขณะเดียวกัน **Excel ยังคงเป็นเครื่องมือทำงานของผู้ใช้งาน**

ไฟล์ที่ระบบสร้างจึงต้องยังคง:

* รูปแบบ Template
* Formula
* Layout
* Header
* Merged Cells
* ข้อมูลที่ผู้ใช้งานต้องแก้ไขต่อใน Excel

---

# 2. Business Flow

## 2.1 Import PO

ผู้ใช้งานสามารถนำเข้า PO เข้าระบบก่อน แล้วนำ PO ที่ Import แล้วไปใช้สร้าง Plan ภายหลังได้

```text
Import PO
   ↓
ตรวจสอบข้อมูล
   ↓
Resolve Location / Product ตามที่จำเป็น
   ↓
PO พร้อมสำหรับสร้าง Plan
```

ระบบเก็บข้อมูล PO ลง Database เป็นหลัก

ไฟล์ต้นฉบับไม่ได้ถูกใช้เป็น Database และไฟล์ชั่วคราวสามารถถูกลบหลังจากนำเข้าข้อมูลแล้ว

---

## 2.2 Create Plan

ระบบรองรับการสร้าง Plan ได้ 2 รูปแบบ

### แบบที่ 1 — Import PO ระหว่างสร้าง Plan

```text
สร้างแผน
   ↓
นำเข้า PO
   ↓
ตรวจสอบข้อมูล
   ↓
กำหนดยอดเผื่อ
   ↓
คำนวณ
   ↓
Production Plan
+
Logistic Plan
```

### แบบที่ 2 — ใช้ PO ที่ Import ไว้แล้ว

```text
สร้างแผน
   ↓
เลือก PO ที่เคย Import
   ↓
กำหนดยอดเผื่อ
   ↓
คำนวณ
   ↓
Production Plan
+
Logistic Plan
```

---

## 2.3 Existing Plan + Add PO

Plan ไม่ได้มีความหมายว่า "สร้างครั้งเดียวแล้วจบ"

ใน Workflow จริง PO สามารถเข้ามาเพิ่มเติมภายหลังได้

ระบบจึงรองรับแนวคิด:

```text
Plan เดิม
   ↓
นำเข้า / เพิ่ม PO
   ↓
รวม PO ชุดใหม่กับ Plan เดิม
   ↓
Recalculate
   ↓
Plan ล่าสุด
```

การเพิ่ม PO ไม่ควรบังคับให้ผู้ใช้งานสร้าง Plan ใหม่ ถ้า Business Flow จริงต้องการให้ใช้ Plan เดิมต่อ

---

# 3. รอบเย็น → รอบเช้า

เป็นหนึ่งใน Workflow สำคัญของระบบ

## 3.1 รอบเย็น

```text
กด "สร้างแผน"
      ↓
นำเข้า PO รอบเย็น
      ↓
ระบบคำนวณ
      ↓
ได้ Plan
      ↓
Download Excel
```

Excel ที่ได้ต้องอ้างอิง Template และยังมี Formula สำหรับให้ผู้ใช้งานตรวจสอบหรือแก้ไขต่อภายหลังได้

---

## 3.2 รอบเช้า

เมื่อมี PO รอบเช้าเข้ามาภายหลัง:

```text
เข้า Plan ล่าสุดของรอบเย็น
      ↓
เพิ่ม PO รอบเช้า
      ↓
รวม PO รอบเย็น + รอบเช้า
      ↓
Recalculate
      ↓
Plan ล่าสุด
      ↓
Download Excel
```

ดังนั้น Plan จึงเป็นข้อมูลที่สามารถ **เปลี่ยนแปลงและคำนวณใหม่ตาม PO ที่เพิ่มเข้ามา** ได้

---

# 4. ยอดเผื่อ (Buffer)

ยอดเผื่อเป็นข้อมูลที่ใช้ในการจัดทำ Plan และสามารถเปลี่ยนแปลงได้ภายหลัง

ระบบรองรับ:

```text
Create Plan
   ↓
กำหนดยอดเผื่อ
   ↓
Calculate
```

และ:

```text
Existing Plan
   ↓
แก้ไขยอดเผื่อ
   ↓
Recalculate
```

การแก้ยอดเผื่อของ Plan เดิมใช้ `plan_run_id` เดิม ไม่สร้าง Plan ใหม่ซ้อนโดยไม่จำเป็น

ค่าเริ่มต้นของยอดเผื่อสามารถอ้างอิงจากค่าที่บันทึกล่าสุดในระบบ

> สูตรหรือหลักเกณฑ์ทางธุรกิจสำหรับการคำนวณยอดเผื่ออัตโนมัติยังไม่ถูกกำหนดเป็น Business Rule ในระบบ

---

# 5. Production Plan และ Logistic Plan

## Production Plan

ใช้สำหรับวางแผนการผลิตจาก PO

ข้อมูลหลักมาจาก:

```text
PO
+
Buffer
+
Production Template
      ↓
Production Plan
```

## Logistic Plan / แพลนกระจาย

ใช้สำหรับจัดข้อมูลสินค้าและจำนวนตามจุดส่ง / Sub-location / PO

โครงสร้าง Logistic Group สามารถตั้งค่าได้ผ่านระบบ โดยไม่จำเป็นต้องแก้ Code เมื่อมี Group เพิ่มขึ้นในกรณีที่มี Template รองรับแล้ว

> **แพลนรถ (Truck Plan) ไม่ใช่ Logistic Plan**
>
> Logistic Plan เป็นข้อมูลสำหรับการกระจายสินค้า
> Truck Plan เป็นขั้นตอนถัดไปสำหรับนำ Logistic Plan ไปจัดกลุ่มเข้ารถ

---

# 6. Location และ Logistic Group

ระบบแยกแนวคิดระหว่าง:

* Location / FC
* Sub-location
* Logistic Group
* Logistic Template

ตัวอย่างโครงสร้างทางธุรกิจ:

```text
Logistic Group
├── บางบัวทอง
│   └── บาร์ระบุวันผลิต
├── มหาชัย
├── สุวรรณ
└── รอบเช้า ต่างจังหวัด
```

ข้อมูล Location Mapping ถูกเก็บใน Database และใช้ประกอบการจัดกลุ่มข้อมูลจริงในการคำนวณ

เมื่อพบ FC ที่ระบบยังไม่รู้จัก ผู้ใช้งานสามารถ Resolve Mapping ผ่าน Web ก่อนสร้าง Plan

---

# 7. Template Management

Template Excel เป็นส่วนสำคัญของระบบ

ระบบรองรับ:

* Template Version
* Template Group
* Group Activation
* Version History
* Upload Template
* Validate Template
* Restore Version
* Download Template Version
* เก็บ Original Filename
* Active Group เป็น Source of Truth ของ Template ที่ใช้งานจริง

Production Template และ Logistic Template สามารถมีโครงสร้างแตกต่างกันตาม Business Requirement

---

# 8. Dynamic Template

ส่วนนี้เป็น **งานพัฒนาหลักที่ยังอยู่ใน Roadmap**

เป้าหมายคือให้ระบบรองรับ Template ที่เปลี่ยนจำนวนข้อมูลได้จริงทั้งสองแกน

### Dynamic Row

จำนวน Product / SKU สามารถเพิ่มหรือลดได้

```text
Product A
Product B
Product C
...
Product N
```

ไม่ควรผูกกับจำนวน Row ที่มีอยู่ใน Template เดิมแบบตายตัว

### Dynamic Column

จำนวน PO สามารถเพิ่มหรือลดได้

```text
PO1 | PO2 | PO3 | ... | PON
```

โดยลำดับ PO ปัจจุบันอ้างอิงจากการเรียงเลข PO

### Dynamic Formula

Formula ต้องสามารถปรับตาม:

* จำนวน Product
* จำนวน PO
* Location
* Sub-location
* Buffer
* Total
* Production Quantity
* Logistic Quantity

### ลด Hardcode

จุดที่ผูกกับตำแหน่ง Excel โดยไม่จำเป็นจะถูกทยอยตรวจสอบและลดลง

อย่างไรก็ตาม **Business Rule ที่ตั้งใจให้เป็นค่าคงที่ไม่ควรถูกลบเพียงเพราะเป็น hardcode**

ตัวอย่างเช่น Business Concept อย่าง `รอบเช้าต่างจังหวัด` ไม่ได้หมายความว่าเป็น Bad Hardcode โดยอัตโนมัติ

---

# 9. Web ↔ Excel

ระบบมีแนวคิด:

> **Database เป็น Source of Truth ของข้อมูล Plan**

ไม่ใช่ให้ Excel เป็น Database

Flow หลัก:

```text
PO / Plan Data
      ↓
Database
      ↓
Calculation
      ↓
PlanSkuResult
      ↓
Web
      ↓
Excel Generation
```

`PlanSkuResult` ทำหน้าที่เป็นข้อมูลผลลัพธ์ระดับ:

```text
1 SKU × 1 Column
```

เพื่อให้ Web สามารถแสดงผลจากข้อมูลที่คำนวณแล้ว โดยไม่จำเป็นต้องอ่านค่าจาก Excel ที่ถูกสร้างขึ้นมาใหม่ทุกครั้ง

ในขณะเดียวกัน เมื่อผู้ใช้งาน Download Excel ระบบสามารถสร้างไฟล์ใหม่จาก:

```text
Database
+
Current Template
+
Formula / Layout
```

---

# 10. Data Model หลัก

ระบบใช้ Django ORM ร่วมกับ PostgreSQL

โมเดลสำคัญในฝั่ง CP All ได้แก่:

* `ProductMaster`
* `LocationMapping`
* `LogisticGroup`
* `PoImport`
* `PoLine`
* `PlanRun`
* `PlanRunLogisticFile`
* `TemplateVersion`
* `TemplateGroup`
* `TemplateGroupItem`
* `PlanSkuResult`

ความสัมพันธ์หลัก:

```text
PO Import
   │
   └── PO Line
          │
          ├── Product / SKU
          │
          └── Location / Sub-location
                    │
                    ▼
                 PlanRun
                    │
          ┌─────────┴─────────┐
          ▼                   ▼
 Production Plan       Logistic Plan
          │                   │
          └─────────┬─────────┘
                    ▼
              PlanSkuResult
```

---

# 11. Data-first Architecture

ระบบใช้แนวคิด **Data-first**

เมื่อ Import PO:

```text
Excel PO
   ↓
Parse
   ↓
Database
   ↓
ไฟล์ต้นฉบับไม่ใช่ Source of Truth
```

เมื่อ Download Plan:

```text
Database
   +
Template Version
   ↓
Generate Excel
   ↓
Download
```

ข้อดีคือ:

* ไม่ต้องเก็บไฟล์ Excel ทุกครั้ง
* สามารถสร้างไฟล์ใหม่จากข้อมูลล่าสุดได้
* Web สามารถอ่านข้อมูลจาก Database
* Template Version ที่ใช้สร้าง Plan สามารถถูกควบคุมได้
* ลดปัญหาไฟล์บน Disk ไม่ตรงกับข้อมูลในระบบ

---

# 12. Inactive SKU

`ProductMaster.is_active=False` ใช้ระบุ SKU ที่ไม่ควรนำมาสร้าง Plan ในสถานะปัจจุบัน

หาก PO มี SKU ที่ Inactive และมีการสั่งจริง:

```text
PO
 ↓
Inactive SKU
 ↓
BLOCK
```

ระบบตรวจสอบก่อนเข้าสู่การสร้าง Plan และมีการตรวจซ้ำใน Flow ที่เกี่ยวข้องกับการคำนวณใหม่

ถ้า SKU Inactive แต่ไม่มี PO ในรอบนั้น ระบบสามารถซ่อน SKU ดังกล่าวจากผลลัพธ์ที่เกี่ยวข้องโดยไม่ลบ Row จริงจาก Template เพื่อไม่ให้ Formula และโครงสร้าง Template เสียหาย

---

# 13. Architecture

โครงสร้างหลักของระบบ:

```text
core/
├── db.py
└── models.py

portal/

customers/
└── cpall/
    ├── logic/
    │   ├── po_parser.py
    │   ├── po_regenerator.py
    │   ├── po_view_data.py
    │   ├── grouping.py
    │   ├── excel_export.py
    │   ├── logistic_plan_export.py
    │   ├── plan_runner.py
    │   ├── plan_regenerator.py
    │   ├── plan_result_extractor.py
    │   ├── plan_view_data.py
    │   ├── template_manager.py
    │   ├── location_mapping_manager.py
    │   ├── product_master_manager.py
    │   ├── config_loader.py
    │   └── date_utils.py
    │
    ├── models.py
    ├── views.py
    ├── urls.py
    ├── admin.py
    │
    ├── config/
    │   ├── sku_master.yaml
    │   └── location_mapping.yaml
    │
    ├── excel_templates/
    │
    └── management/
        └── commands/

sql/
└── schema.sql

webproject/
```

แนวคิดของ Architecture คือ:

> **1 Customer = 1 Django App**

ปัจจุบันมี:

```text
cpall = CP All / 7-11
```

และโครงสร้างถูกเตรียมไว้เพื่อรองรับ Customer อื่นในอนาคต

---

# 14. PostgreSQL Row-Level Security

ระบบเตรียมการแยกข้อมูลระหว่าง Customer ด้วย PostgreSQL RLS

มีการแยก Connection Role:

### `app_role`

ใช้สำหรับการทำงานปกติของ Application

```text
Application
   ↓
app_role
   ↓
PostgreSQL RLS
```

### `postgres`

ใช้สำหรับงานระดับ Database เช่น Schema / Migration เท่านั้น

ไม่ควรใช้เป็น Runtime Connection สำหรับ Query ข้อมูลทั่วไป เพราะ Superuser สามารถ bypass RLS ได้

---

# 15. Technology Stack

* Python
* Django
* PostgreSQL
* Django ORM
* HTMX
* OpenPyXL
* LibreOffice สำหรับตรวจสอบ / คำนวณ Formula ในกระบวนการสร้างผลลัพธ์
* Docker / Docker Compose
* VS Code Dev Containers
* Django Admin
* Kubi Admin Theme

---

# 16. Development Setup

## Requirements

* Docker Desktop
* VS Code
* VS Code Dev Containers Extension

## Start

เปิด Project ใน VS Code แล้วเลือก:

```text
Dev Containers: Reopen in Container
```

หลังจาก Container พร้อมแล้ว:

```bash
python manage.py runserver 0.0.0.0:8000
```

เปิด:

```text
http://localhost:8000/
```

สำหรับ Django Admin:

```bash
DB_USER=postgres DB_PASSWORD=postgres python manage.py createsuperuser
```

---

# 17. LAN Access

รัน Django ด้วย:

```bash
python manage.py runserver 0.0.0.0:8000
```

และกำหนด `CSRF_TRUSTED_ORIGINS` ใน `.env` ให้ครอบคลุม IP ของเครื่องที่ให้บริการ

ตัวอย่าง:

```env
CSRF_TRUSTED_ORIGINS=http://localhost:8000,http://127.0.0.1:8000,http://192.168.x.x:8000
```

จากเครื่องอื่นใน Network เดียวกัน:

```text
http://<SERVER-IP>:8000
```

---

# 18. Code Quality

สามารถใช้ Ruff ตรวจสอบ Code:

```bash
ruff check .
```

และแก้สิ่งที่ Ruff แก้ได้อัตโนมัติ:

```bash
ruff check . --fix
```

---

# 19. Current Development Status

## ✅ Completed / Foundation

* Portal
* Home page
* PO Import
* PO Data-first storage
* Location Mapping
* Product Master
* Production Plan
* Logistic Plan
* Buffer input / edit
* Plan recalculation foundation
* Plan result storage
* Template Management
* Template Version
* Template Group
* Logistic Group configuration
* Active Template Group
* Template Version History
* Web Plan Detail
* Excel regeneration / download
* ZIP download
* Inactive SKU validation
* LAN development access
* PostgreSQL / RLS foundation

---

# 20. Current Roadmap

## 🔴 Priority 1 — Production Plan Calculation & Recalculation

ตรวจสอบและทำให้ Business Flow สมบูรณ์:

* Create Plan + Import PO
* Create Plan + Existing PO
* Existing Plan + Add PO
* รอบเย็น → สร้าง Plan
* รอบเช้า → เพิ่ม PO เข้า Plan เดิม
* Edit Buffer
* Recalculate
* PlanSkuResult
* Web Plan Detail
* PO Set Control
* Calculation Source of Truth

---

## 🔴 Priority 2 — Dynamic Template

ทำให้ Excel รองรับข้อมูลที่เปลี่ยนแปลงจริง:

* Dynamic Product Rows
* Dynamic PO Columns
* Dynamic Location / Sub-location
* Dynamic Formula
* ลด Hardcode ที่ไม่จำเป็น
* Preserve Template Layout
* Production Excel
* Logistic Excel
* Web ↔ Excel Consistency

---

## 🟠 Priority 3 — Test / UAT

ทดสอบ End-to-End หลัง Business Flow และ Dynamic Template มีความนิ่ง

ครอบคลุม:

* PO
* Product
* Buffer
* Recalculate
* รอบเย็น
* รอบเช้า
* Multiple PO
* Multiple Location
* Multiple Sub-location
* Dynamic Row
* Dynamic Column
* Dynamic Formula
* Web = Excel

---

## 🟡 Priority 4 — Auth / Permission

* Login
* Role
* Admin / User
* Template Permission
* PO Permission
* Plan Permission

---

## 🟡 Priority 5 — LAN / Deployment

* Production configuration
* `DEBUG=False`
* CSRF
* Database configuration
* Backup
* Deployment

---

## 🟡 Priority 6 — Truck Plan

เป็น Feature ใหม่ที่ต่อยอดจาก Logistic Plan

```text
Logistic Plan
      ↓
Destination
      ↓
Product + Quantity
      ↓
Vehicle Capacity
      ↓
แนะนำจัดเข้ารถ
```

เป้าหมายแรกคือ **แนะนำการจัดสินค้าเข้ารถ** ไม่จำเป็นต้องเริ่มจากการกำหนดทะเบียนรถ

งานนี้จะทำหลัง Core Production / Logistic Flow มีความนิ่งแล้ว

---

## 🟣 Priority 7 — Technical Debt

งานกลุ่มนี้จะทำแบบเลือกเป็นจุด ๆ และต้องตรวจ Cross-file Dependency ก่อนทุกครั้ง

รายการที่อยู่ในกลุ่มนี้ เช่น:

* LocationMapping ↔ LogisticGroup ↔ Template consistency
* Log / Audit Trail
* Filesystem / Transaction consistency
* Customer consistency
* `CREATE_PRODUCT` metadata
* Legacy compatibility cleanup
* Refactor

**Legacy code จะไม่ถูกลบเพียงเพราะดูเก่า**

ก่อนลบต้องตรวจว่า:

1. มี Reference หรือไม่
2. มี Active Code Path หรือไม่
3. มี Template / Data ที่พึ่งพาหรือไม่
4. มี Test หรือไม่
5. กระทบ Production / Logistic หรือไม่

ถ้ายังมีความไม่แน่นอน จะไม่ลบ

---

# 21. Design Principles

## 21.1 Workflow First

ระบบต้องเข้าใจ Workflow จริงของผู้ใช้งานก่อนออกแบบ Feature

```text
Business Workflow
       ↓
System Support
```

ไม่ใช่:

```text
System Limitation
       ↓
Force User Workflow
```

---

## 21.2 Business Rule ≠ Technical Limitation

สิ่งที่ระบบต้อง Block ควรเป็นสิ่งที่จำเป็น เช่น:

* Data Integrity
* Business Rule
* ข้อมูลที่ไม่เพียงพอต่อการคำนวณ
* ข้อมูลที่ขัดกับเงื่อนไขของระบบ

แต่ไม่ควร Block เพียงเพราะ:

> "Code ปัจจุบันทำได้แค่นี้"

---

## 21.3 Excel ยังคงเป็นเครื่องมือของ User

ระบบไม่ได้พยายามกำจัด Excel

เป้าหมายคือ:

```text
ระบบช่วยเตรียมข้อมูล
       ↓
ระบบช่วยคำนวณ
       ↓
ระบบสร้าง Excel
       ↓
User ตรวจ / แก้ / ใช้งานต่อ
```

---

## 21.4 Database เป็น Source of Truth

Excel เป็น Output / Working File

ไม่ใช่ Database

```text
Database
   ↓
Calculation
   ↓
Web
   ↓
Excel
```

---

## 21.5 ตรวจ Impact ก่อนแก้

ทุกการแก้ไข โดยเฉพาะ Shared Function / Shared Template / Exporter ต้องตรวจ:

* ใครเรียกใช้
* ถูกใช้ใน Flow ไหน
* มี Dependency กับไฟล์ไหน
* มีผลกับ Production หรือ Logistic หรือไม่
* มีผลกับ Template Version หรือไม่
* มีผลกับ Web หรือ Excel หรือไม่

**ไม่แก้แบบ isolated โดยไม่ตรวจ Cross-file Dependency**

---

# 22. Important Development Lessons

## Schema Migration

เมื่อแก้ `sql/schema.sql` ต้องทดสอบอย่างน้อย 3 กรณี:

1. Database ที่มีข้อมูลอยู่แล้ว
2. Run Schema ซ้ำกับ Database เดิม
3. Fresh Database

เพราะ `schema.sql` ต้องรองรับทั้ง migration และ fresh installation

---

## Excel Cell Clearing

ใน OpenPyXL:

```python
ws.cell(row, col).value = None
```

ใช้สำหรับเคลียร์ค่าที่มีอยู่จริง

ไม่ควรตีความ `cell(value=None)` ว่าเป็นการเคลียร์ค่าเสมอไป

---

## End-to-End Testing

Feature ที่ผ่านการทดสอบแยกกันไม่ได้แปลว่า Full Workflow จะไม่มี Bug

ควรมีการทดสอบแบบ:

```text
Import PO
   ↓
Duplicate Check
   ↓
Location Resolve
   ↓
Product Resolve
   ↓
Buffer
   ↓
Create Plan
   ↓
View Plan
   ↓
Edit Buffer
   ↓
Recalculate
   ↓
Download Excel
```

---

# 23. Project Philosophy

โปรเจกต์นี้ไม่ได้มีเป้าหมายเพียง:

> "สร้าง Excel อัตโนมัติ"

แต่มีเป้าหมายเพื่อ:

> **ลดงาน Manual ที่ไม่จำเป็น โดยให้ระบบเข้ามาช่วยใน Workflow ที่ผู้ใช้งานทำอยู่แล้ว**

ผู้ใช้งานยังเป็นผู้ตัดสินใจ

ระบบทำหน้าที่:

* จัดข้อมูล
* ตรวจสอบ
* คำนวณ
* สร้างผลลัพธ์
* ลดงานซ้ำ
* ลดความผิดพลาด
* ทำให้ข้อมูลตรวจสอบย้อนหลังได้ง่ายขึ้น

และเมื่อ Business Workflow เปลี่ยน ระบบควรสามารถปรับตาม Workflow ได้โดยไม่สร้างข้อจำกัดใหม่ให้ผู้ใช้งานโดยไม่จำเป็น
