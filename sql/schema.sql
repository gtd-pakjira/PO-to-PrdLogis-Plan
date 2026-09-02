-- schema.sql
-- แบ่งเป็น 2 ส่วน: "Core" (ของกลาง ใช้ร่วมกันทุกลูกค้า) กับ "CPAll" (ข้อมูลเฉพาะลูกค้า 7-11)
-- ตารางฝั่ง CPAll ทุกตารางมี customer_id + Row-Level Security (RLS) กันข้อมูลรั่วข้ามลูกค้า
-- แม้โค้ดจะลืม filter customer_id เอง ฐานข้อมูลก็จะกันให้อีกชั้นหนึ่ง

-- ============================================================
-- CORE — ของกลาง ใช้ร่วมกันทุกลูกค้า (ไม่มี RLS เพราะตัวมันเองไม่ใช่ข้อมูลอ่อนไหว)
-- ============================================================

CREATE TABLE IF NOT EXISTS customer (
    id              SERIAL PRIMARY KEY,
    code            VARCHAR(20) UNIQUE NOT NULL,   -- เช่น 'cpall' — ใช้อ้างอิงในโค้ด (คงที่ ไม่เปลี่ยน)
    name_th         TEXT NOT NULL,                 -- เช่น 'CP All (7-11)' — ใช้แสดงผลในหน้าเว็บ
    created_at      TIMESTAMP DEFAULT now()
);

-- seed ลูกค้าแรก (7-11) — ทำครั้งเดียว ไม่ทับถ้ามีอยู่แล้ว
INSERT INTO customer (code, name_th) VALUES ('cpall', 'CP All (7-11)')
    ON CONFLICT (code) DO NOTHING;

-- ============================================================
-- CPALL — ข้อมูลเฉพาะลูกค้า 7-11 ทั้งหมด (RLS เปิดทุกตาราง)
-- ============================================================

-- ---------- Config tables (โหลดจาก YAML ตอน seed / sync) ----------

CREATE TABLE IF NOT EXISTS sku_master (
    customer_id     INTEGER NOT NULL REFERENCES customer(id),
    barcode         VARCHAR(20) NOT NULL,
    name_th         TEXT NOT NULL,
    name_en         TEXT,
    pack_size       INTEGER NOT NULL,      -- บรรจุ/ตก. (ชิ้นต่อลัง)
    unit_price      NUMERIC(10,2),         -- ราคาล่าสุด (อ้างอิง — ราคาจริงต่อรอบมาจาก PO)
    note            TEXT,
    updated_at      TIMESTAMP DEFAULT now(),
    PRIMARY KEY (customer_id, barcode)     -- บาร์โค้ดซ้ำกันได้ข้ามลูกค้า (คนละลูกค้า = คนละสินค้า)
);

CREATE TABLE IF NOT EXISTS location_mapping (
    customer_id     INTEGER NOT NULL REFERENCES customer(id),
    fc_code         VARCHAR(10) NOT NULL,      -- Delivery Location Number จาก PO เช่น FC08
    name_th         TEXT NOT NULL,
    "group"         VARCHAR(50) NOT NULL,      -- บางบัวทอง / มหาชัย / สุวรรณภูมิ
    sub_location    VARCHAR(50),
    updated_at      TIMESTAMP DEFAULT now(),
    PRIMARY KEY (customer_id, fc_code)     -- รหัสสถานที่ซ้ำกันได้ข้ามลูกค้าเช่นกัน
);

-- ---------- PO Import (raw data ต่อรอบ) ----------

CREATE TABLE IF NOT EXISTS po_import (
    id              SERIAL PRIMARY KEY,
    customer_id     INTEGER NOT NULL REFERENCES customer(id),
    source_filename TEXT NOT NULL,
    imported_at     TIMESTAMP DEFAULT now(),
    imported_by     VARCHAR(100),
    production_date DATE,   -- วันที่ผลิต (เดิมชื่อ po_date — ตั้งชื่อผิดความหมายมาตั้งแต่แรก แก้แล้ว)
    po_date         DATE,   -- วันที่ PO (เดิมชื่อ delivery_date — ก็ตั้งชื่อผิดความหมายเหมือนกัน)
    total_rows      INTEGER,
    status          VARCHAR(20) DEFAULT 'imported'
);

CREATE TABLE IF NOT EXISTS po_line (
    id                  SERIAL PRIMARY KEY,
    po_import_id        INTEGER NOT NULL REFERENCES po_import(id) ON DELETE CASCADE,
    po_number           VARCHAR(30) NOT NULL,
    po_date             DATE,
    delivery_date       DATE,
    delivery_time       VARCHAR(10),
    fc_code             VARCHAR(10) NOT NULL,
    delivery_location    TEXT,
    line_no             INTEGER,
    barcode             VARCHAR(20) NOT NULL,
    item_name           TEXT,
    qty_ordered         NUMERIC(10,2) NOT NULL,
    unit_type           VARCHAR(10),
    net_case_price      NUMERIC(10,2),
    total_amount        NUMERIC(12,2)
    -- หมายเหตุ: ตารางนี้ไม่มี customer_id ของตัวเอง (อ้างอิงผ่าน po_import_id พอ) และไม่เปิด RLS
    -- ตรงๆ — โค้ดของ cpall เข้าถึงผ่าน po_import_id ที่ผ่านการเช็คสิทธิ์จาก po_import แล้วเสมอ
);

CREATE INDEX IF NOT EXISTS idx_po_line_import ON po_line(po_import_id);
CREATE INDEX IF NOT EXISTS idx_po_line_fc ON po_line(fc_code);
CREATE INDEX IF NOT EXISTS idx_po_line_barcode ON po_line(barcode);

-- ---------- Plan Run (ประวัติการสร้างแผน — ใช้โดยหน้าเว็บ) ----------

CREATE TABLE IF NOT EXISTS plan_run (
    id                      SERIAL PRIMARY KEY,
    customer_id             INTEGER NOT NULL REFERENCES customer(id),
    created_at              TIMESTAMP DEFAULT now(),
    output_dir              TEXT NOT NULL,
    production_plan_path    TEXT,
    production_plan_status  VARCHAR(20) DEFAULT 'success',
    production_plan_error   TEXT
);

CREATE TABLE IF NOT EXISTS plan_run_import (
    plan_run_id     INTEGER NOT NULL REFERENCES plan_run(id) ON DELETE CASCADE,
    po_import_id    INTEGER NOT NULL REFERENCES po_import(id),
    PRIMARY KEY (plan_run_id, po_import_id)
);

CREATE TABLE IF NOT EXISTS plan_run_logistic_file (
    id              SERIAL PRIMARY KEY,
    plan_run_id     INTEGER NOT NULL REFERENCES plan_run(id) ON DELETE CASCADE,
    group_name      VARCHAR(50) NOT NULL,
    status          VARCHAR(20) NOT NULL,
    file_path       TEXT,
    error_message   TEXT
);

-- ---------- Template Versioning (Phase 1.6 sub-phase 1) ----------
-- เก็บทุกเวอร์ชันของไฟล์เทมเพลตที่เคยอัปโหลดถาวร (ต่างจากของเดิมที่เก็บ backup แค่ 1 ชั้น) —
-- ให้กู้คืนไปเวอร์ชันไหนก็ได้ในประวัติ และให้แผนแต่ละแผนอ้างอิงได้ว่าตอนสร้างใช้เทมเพลตเวอร์ชันไหน
-- (ดาวน์โหลดแผนเก่าซ้ำ จะได้ตัวเลขตรงกับตอนสร้างแผนเป๊ะ แม้เทมเพลตปัจจุบันจะถูกแก้ไปแล้วก็ตาม)
CREATE TABLE IF NOT EXISTS template_version (
    id                  SERIAL PRIMARY KEY,
    customer_id         INTEGER NOT NULL REFERENCES customer(id),
    template_key        VARCHAR(50) NOT NULL,   -- 'production_plan', 'logistic_บางบัวทอง' ฯลฯ
    version_number      INTEGER NOT NULL,
    file_path           TEXT NOT NULL,          -- ที่เก็บถาวรของไฟล์เวอร์ชันนี้ (คนละที่กับไฟล์ live)
    is_active           BOOLEAN NOT NULL DEFAULT FALSE,
    uploaded_at         TIMESTAMP DEFAULT now(),
    validation_summary  TEXT,                   -- เช่น "sku_count=19" เก็บไว้ดูย้อนหลังเฉยๆ
    UNIQUE (customer_id, template_key, version_number)
);
CREATE INDEX IF NOT EXISTS idx_template_version_key ON template_version(customer_id, template_key);

-- เชื่อมแผนที่สร้างไว้เข้ากับเทมเพลตเวอร์ชันที่ใช้จริงตอนนั้น — NULL ได้สำหรับแผนเก่าก่อนมีระบบนี้
-- (การ "บันทึกจริง" ว่าใช้เวอร์ชันไหนทำใน sub-phase 3 ตอนเปลี่ยน flow สร้างแผน — คอลัมน์นี้แค่เตรียมที่ไว้)
ALTER TABLE plan_run ADD COLUMN IF NOT EXISTS production_template_version_id INTEGER
    REFERENCES template_version(id);
ALTER TABLE plan_run_logistic_file ADD COLUMN IF NOT EXISTS template_version_id INTEGER
    REFERENCES template_version(id);

CREATE INDEX IF NOT EXISTS idx_plan_run_import_run ON plan_run_import(plan_run_id);
CREATE INDEX IF NOT EXISTS idx_plan_run_logistic_run ON plan_run_logistic_file(plan_run_id);

-- ---------- ผลลัพธ์ต่อ SKU/คอลัมน์ (Phase 1.6 sub-phase 2) ----------
-- "1 แถว = 1 SKU x 1 คอลัมน์" (เช่น "บางบัวทอง" หรือ "ชลบุรี PO2") — เก็บยอดสั่งจริง (ไม่ใช่จากสูตร
-- จำลอง) และค่าที่ LibreOffice คำนวณจากสูตรจริงในเทมเพลต (pack_text/return_qty/basket_total)
-- แยกจากไฟล์ Excel — สร้างแผนแล้วอ่านค่าจากตารางนี้แสดงเว็บได้เลย ไม่ต้องพึ่งไฟล์ที่อาจหายไปได้
-- (denormalize grand_total/buffer_qty/return_qty/basket_total ซ้ำทุกแถวของ SKU เดียวกันโดยตั้งใจ —
-- ค่าพวกนี้เป็นระดับ "ต่อ SKU" ไม่ใช่ "ต่อคอลัมน์" แต่ยอมซ้ำเพื่อไม่ต้อง JOIN ตอนอ่านแสดงผล)
CREATE TABLE IF NOT EXISTS plan_sku_result (
    id              SERIAL PRIMARY KEY,
    plan_run_id     INTEGER NOT NULL REFERENCES plan_run(id) ON DELETE CASCADE,
    sheet_type      VARCHAR(20) NOT NULL,       -- 'production' | 'logistic'
    group_name      VARCHAR(50),                -- NULL สำหรับ production, ชื่อกลุ่มสำหรับ logistic
    barcode         VARCHAR(20) NOT NULL,
    name_th         TEXT,
    name_en         TEXT,
    price           NUMERIC(10,2),               -- เฉพาะ production
    pack_size       INTEGER,
    column_label    VARCHAR(100) NOT NULL,       -- เช่น "บางบัวทอง" (production) หรือ "ชลบุรี PO2" (logistic)
    qty             NUMERIC(10,2),
    pack_text       VARCHAR(50),                 -- ผลจาก LibreOffice จริง เช่น "3 + 25 P"
    grand_total     NUMERIC(12,2),
    buffer_qty      NUMERIC(10,2),               -- เฉพาะ production
    return_qty      NUMERIC(10,2),               -- เฉพาะ production, ผลจาก LibreOffice จริง
    basket_total    INTEGER                      -- เฉพาะ logistic, ผลจาก LibreOffice จริง
);
CREATE INDEX IF NOT EXISTS idx_plan_sku_result_run ON plan_sku_result(plan_run_id);
CREATE INDEX IF NOT EXISTS idx_plan_sku_result_sheet ON plan_sku_result(plan_run_id, sheet_type, group_name);

-- ============================================================
-- MIGRATION — สำหรับฐานข้อมูลที่เคยรัน schema.sql เวอร์ชันเก่ามาก่อน (มี sku_master/location_mapping/
-- po_import/plan_run อยู่แล้วแบบไม่มี customer_id) ย้ายเข้าโครงสร้างใหม่แบบไม่ทำข้อมูลเดิมหาย
-- รันซ้ำได้เรื่อยๆ ปลอดภัย (เช็คก่อนทุกขั้นตอน)
-- ============================================================
DO $$
DECLARE
    cpall_id INTEGER;
BEGIN
    SELECT id INTO cpall_id FROM customer WHERE code = 'cpall';

    -- เก็บกวาดตารางที่ไม่ได้ใช้แล้วจากดีไซน์รุ่นก่อนหน้า (buffer_qty ถูกแทนที่ด้วยฟอร์มกรอกยอดเผื่อ
    -- ผ่านเว็บ, production_plan_run/production_plan_line เป็น Module 3 เดิมที่เลิกใช้ไปแล้ว) —
    -- ตารางพวกนี้ค้าง foreign key ชี้มาที่ sku_master แบบ PK เดิม ต้องลบก่อนถึงจะเปลี่ยน PK ได้
    DROP TABLE IF EXISTS buffer_qty;
    DROP TABLE IF EXISTS production_plan_line;
    DROP TABLE IF EXISTS production_plan_run;

    -- sku_master: เพิ่ม customer_id, ใส่ค่าเดิมทั้งหมดเป็นของ cpall, เปลี่ยน PK
    IF EXISTS (SELECT 1 FROM information_schema.tables WHERE table_name = 'sku_master') THEN
        IF NOT EXISTS (SELECT 1 FROM information_schema.columns
                        WHERE table_name = 'sku_master' AND column_name = 'customer_id') THEN
            ALTER TABLE sku_master ADD COLUMN customer_id INTEGER REFERENCES customer(id);
            UPDATE sku_master SET customer_id = cpall_id WHERE customer_id IS NULL;
            ALTER TABLE sku_master ALTER COLUMN customer_id SET NOT NULL;
            ALTER TABLE sku_master DROP CONSTRAINT IF EXISTS sku_master_pkey;
            ALTER TABLE sku_master ADD PRIMARY KEY (customer_id, barcode);
        END IF;
    END IF;

    -- location_mapping: เหมือนกัน
    IF EXISTS (SELECT 1 FROM information_schema.tables WHERE table_name = 'location_mapping') THEN
        IF NOT EXISTS (SELECT 1 FROM information_schema.columns
                        WHERE table_name = 'location_mapping' AND column_name = 'customer_id') THEN
            ALTER TABLE location_mapping ADD COLUMN customer_id INTEGER REFERENCES customer(id);
            UPDATE location_mapping SET customer_id = cpall_id WHERE customer_id IS NULL;
            ALTER TABLE location_mapping ALTER COLUMN customer_id SET NOT NULL;
            ALTER TABLE location_mapping DROP CONSTRAINT IF EXISTS location_mapping_pkey;
            ALTER TABLE location_mapping ADD PRIMARY KEY (customer_id, fc_code);
        END IF;
    END IF;

    -- po_import: แค่เพิ่ม customer_id (PK เดิม (id) ไม่ต้องเปลี่ยน)
    IF EXISTS (SELECT 1 FROM information_schema.tables WHERE table_name = 'po_import') THEN
        IF NOT EXISTS (SELECT 1 FROM information_schema.columns
                        WHERE table_name = 'po_import' AND column_name = 'customer_id') THEN
            ALTER TABLE po_import ADD COLUMN customer_id INTEGER REFERENCES customer(id);
            UPDATE po_import SET customer_id = cpall_id WHERE customer_id IS NULL;
            ALTER TABLE po_import ALTER COLUMN customer_id SET NOT NULL;
        END IF;

        -- แก้ชื่อคอลัมน์ po_date/delivery_date ที่ตั้งผิดความหมายมาตั้งแต่แรก (ดูหมายเหตุบนตาราง) —
        -- แค่เปลี่ยนชื่อ ไม่แตะข้อมูล (ค่าที่กรอกไว้ถูกต้องอยู่แล้ว ผิดแค่ชื่อที่เรียก)
        IF EXISTS (SELECT 1 FROM information_schema.columns
                   WHERE table_name = 'po_import' AND column_name = 'delivery_date') THEN
            -- ต้องย้าย po_date (ชื่อเดิม) ออกจากทางก่อน ไม่งั้นจะมี 2 คอลัมน์ชื่อ po_date ชนกันตอน
            -- rename delivery_date -> po_date
            ALTER TABLE po_import RENAME COLUMN po_date TO production_date;
            ALTER TABLE po_import RENAME COLUMN delivery_date TO po_date;
        END IF;
    END IF;

    -- plan_run: แค่เพิ่ม customer_id เหมือนกัน
    IF EXISTS (SELECT 1 FROM information_schema.tables WHERE table_name = 'plan_run') THEN
        IF NOT EXISTS (SELECT 1 FROM information_schema.columns
                        WHERE table_name = 'plan_run' AND column_name = 'customer_id') THEN
            ALTER TABLE plan_run ADD COLUMN customer_id INTEGER REFERENCES customer(id);
            UPDATE plan_run SET customer_id = cpall_id WHERE customer_id IS NULL;
            ALTER TABLE plan_run ALTER COLUMN customer_id SET NOT NULL;
        END IF;
    END IF;
END $$;

-- ============================================================
-- Row-Level Security — เปิดใช้บนทุกตารางที่มี customer_id ตรงๆ
-- ============================================================
-- หลักการ: แอปตั้งค่า session variable "app.current_customer_id" ทันทีหลังเชื่อมต่อ DB
-- (ดู core/db.py) — Postgres จะกรองให้อัตโนมัติว่าเห็นได้แค่แถวของลูกค้านั้น ต่อให้โค้ด SELECT/UPDATE/
-- DELETE ลืม WHERE customer_id = ... ก็ยังปลอดภัย เพราะ FORCE ROW LEVEL SECURITY บังคับด้วยแม้แต่
-- เจ้าของตาราง/superuser ก็ไม่ยกเว้น (ปกติ Postgres จะข้าม RLS ให้ table owner โดย default)

ALTER TABLE sku_master ENABLE ROW LEVEL SECURITY;
ALTER TABLE sku_master FORCE ROW LEVEL SECURITY;
DROP POLICY IF EXISTS customer_isolation ON sku_master;
CREATE POLICY customer_isolation ON sku_master
    USING (customer_id::text = current_setting('app.current_customer_id', true))
    WITH CHECK (customer_id::text = current_setting('app.current_customer_id', true));

ALTER TABLE location_mapping ENABLE ROW LEVEL SECURITY;
ALTER TABLE location_mapping FORCE ROW LEVEL SECURITY;
DROP POLICY IF EXISTS customer_isolation ON location_mapping;
CREATE POLICY customer_isolation ON location_mapping
    USING (customer_id::text = current_setting('app.current_customer_id', true))
    WITH CHECK (customer_id::text = current_setting('app.current_customer_id', true));

ALTER TABLE po_import ENABLE ROW LEVEL SECURITY;
ALTER TABLE po_import FORCE ROW LEVEL SECURITY;
DROP POLICY IF EXISTS customer_isolation ON po_import;
CREATE POLICY customer_isolation ON po_import
    USING (customer_id::text = current_setting('app.current_customer_id', true))
    WITH CHECK (customer_id::text = current_setting('app.current_customer_id', true));

ALTER TABLE plan_run ENABLE ROW LEVEL SECURITY;
ALTER TABLE plan_run FORCE ROW LEVEL SECURITY;
DROP POLICY IF EXISTS customer_isolation ON plan_run;
CREATE POLICY customer_isolation ON plan_run
    USING (customer_id::text = current_setting('app.current_customer_id', true))
    WITH CHECK (customer_id::text = current_setting('app.current_customer_id', true));

ALTER TABLE template_version ENABLE ROW LEVEL SECURITY;
ALTER TABLE template_version FORCE ROW LEVEL SECURITY;
DROP POLICY IF EXISTS customer_isolation ON template_version;
CREATE POLICY customer_isolation ON template_version
    USING (customer_id::text = current_setting('app.current_customer_id', true))
    WITH CHECK (customer_id::text = current_setting('app.current_customer_id', true));

-- หมายเหตุขอบเขต: po_line / plan_run_import / plan_run_logistic_file ไม่มี customer_id ของตัวเอง
-- (อ้างอิงผ่านตารางแม่ที่มี RLS อยู่แล้ว) เข้าถึงได้เฉพาะผ่าน po_import_id/plan_run_id ที่โค้ด
-- ดึงมาจากตารางแม่ที่ผ่าน RLS กรองแล้วเท่านั้น — ยังไม่ได้ใส่ RLS ตรงบนตารางลูกพวกนี้เอง (ทำเพิ่มได้
-- ทีหลังด้วย policy แบบ EXISTS subquery ถ้าพบว่าจำเป็น)

-- ============================================================
-- App role — สำคัญมาก: RLS ไม่มีผลกับ superuser (postgres) เลย ไม่ว่าจะตั้ง FORCE หรือไม่ก็ตาม
-- ต้องสร้าง role ใหม่ที่ไม่ใช่ superuser ให้แอปใช้เชื่อมต่อจริง ถึงจะโดน RLS บังคับ — 'postgres'
-- (superuser) เก็บไว้ใช้แค่ตอนรัน schema.sql/migration เท่านั้น (ผ่าน core.db.get_admin_connection())
-- ============================================================

DO $$
BEGIN
    IF NOT EXISTS (SELECT 1 FROM pg_roles WHERE rolname = 'app_role') THEN
        CREATE ROLE app_role LOGIN PASSWORD 'app_password';
    END IF;
END $$;

GRANT SELECT, INSERT, UPDATE, DELETE ON ALL TABLES IN SCHEMA public TO app_role;
GRANT USAGE, SELECT ON ALL SEQUENCES IN SCHEMA public TO app_role;
-- ครอบคลุมทั้งตารางของเราเอง (customer, sku_master, ...) และตารางที่ Django สร้างเอง (auth_user,
-- django_session, ฯลฯ) — Django migrate ต้องรันด้วย DB_ADMIN_USER (superuser) เสมอ เพราะ app_role
-- ไม่มีสิทธิ์ CREATE TABLE ตั้งใจไว้แบบนั้น (กัน RLS/schema ถูกแก้จากจุดที่ไม่ควรแก้) — grant บรรทัดนี้
-- แค่ให้ app_role "ใช้งาน" ตารางที่มีอยู่แล้วได้ ไม่ได้ให้สิทธิ์สร้าง/แก้โครงสร้างตาราง
