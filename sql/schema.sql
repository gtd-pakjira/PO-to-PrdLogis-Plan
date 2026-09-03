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

-- เปลี่ยนชื่อตาราง sku_master -> product_master ("SKU" เป็นศัพท์ฝั่งคลังสินค้า ไม่ตรงกับที่ใช้จริง —
-- เปลี่ยน UI เป็น "รหัสสินค้า/Product Code" ไปแล้ว เปลี่ยนชื่อ table ให้ตรงกันด้วย) — ต้องทำเป็นก้าวแรก
-- สุด (ก่อน migration/CREATE TABLE อื่นใดทั้งหมด) — migration อื่นๆ ที่เหลือทั้งหมดในไฟล์นี้ (ancient
-- migration, surrogate key, is_active) เขียนให้ใช้ชื่อ product_master เสมอ โดยยึดว่าขั้นตอนนี้ทำไปแล้ว
DO $$
BEGIN
    IF EXISTS (SELECT 1 FROM information_schema.tables WHERE table_name = 'sku_master')
       AND NOT EXISTS (SELECT 1 FROM information_schema.tables WHERE table_name = 'product_master') THEN
        ALTER TABLE sku_master RENAME TO product_master;
    END IF;
END $$;

-- ============================================================
-- CPALL — ข้อมูลเฉพาะลูกค้า 7-11 ทั้งหมด (RLS เปิดทุกตาราง)
-- ============================================================

-- ---------- Config tables (โหลดจาก YAML ตอน seed / sync) ----------

-- ถ้ามีตาราง sku_master เก่าอยู่แล้ว (ยังไม่เคย rename) อย่าเพิ่งสร้าง product_master ใหม่ตรงนี้ —
-- ปล่อยให้ migration ท้ายไฟล์ (หลัง ancient migration ที่ยังต้องทำงานกับชื่อ sku_master เดิมก่อน)
-- จัดการ RENAME ให้ก่อน ไม่งั้นจะได้ตารางว่างเปล่าซ้อนกับตารางเก่าที่มีข้อมูลจริง (เจอบั๊กนี้มาแล้ว)
DO $$
BEGIN
    IF NOT EXISTS (SELECT 1 FROM information_schema.tables WHERE table_name = 'product_master')
       AND NOT EXISTS (SELECT 1 FROM information_schema.tables WHERE table_name = 'sku_master') THEN
        CREATE TABLE product_master (
            id              SERIAL PRIMARY KEY,
            customer_id     INTEGER NOT NULL REFERENCES customer(id),
            barcode         VARCHAR(20) NOT NULL,
            name_th         TEXT NOT NULL,
            name_en         TEXT,
            pack_size       INTEGER NOT NULL,
            unit_price      NUMERIC(10,2),
            note            TEXT,
            is_active       BOOLEAN NOT NULL DEFAULT TRUE,
            updated_at      TIMESTAMP DEFAULT now(),
            UNIQUE (customer_id, barcode)
        );
    END IF;
END $$;

CREATE TABLE IF NOT EXISTS location_mapping (
    id              SERIAL PRIMARY KEY,        -- surrogate key (Django Admin ใช้ composite PK ไม่ได้)
    customer_id     INTEGER NOT NULL REFERENCES customer(id),
    fc_code         VARCHAR(10) NOT NULL,      -- Delivery Location Number จาก PO เช่น FC08
    name_th         TEXT NOT NULL,
    "group"         VARCHAR(50) NOT NULL,      -- บางบัวทอง / มหาชัย / สุวรรณภูมิ
    sub_location    VARCHAR(50),
    is_active       BOOLEAN NOT NULL DEFAULT TRUE,  -- ปิดใช้งานจุดส่งที่เลิกใช้แล้วได้ ไม่ต้องลบทิ้ง
    updated_at      TIMESTAMP DEFAULT now(),
    UNIQUE (customer_id, fc_code)          -- รหัสสถานที่ซ้ำกันได้ข้ามลูกค้าเช่นกัน
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
    status          VARCHAR(20) DEFAULT 'imported',
    column_order    JSONB   -- ลำดับ+ชื่อคอลัมน์ทั้งหมดในไฟล์ต้นฉบับ (array) เผื่อชื่อซ้ำกัน (ไฟล์จริง
                             -- มีคอลัมน์ชื่อซ้ำ เช่น "Discount Percentage 1" ปรากฏ 2 รอบ) ใช้ตำแหน่ง
                             -- ไม่ใช่ชื่อจับคู่กับ po_line.all_values ตอนสร้างไฟล์ใหม่ กันปัญหาคอลัมน์ชนกัน
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
    total_amount        NUMERIC(12,2),
    all_values          JSONB   -- ค่าทุกคอลัมน์ของแถวนี้ตามลำดับเดียวกับ po_import.column_order (array)
                                 -- เก็บไว้ให้สร้างไฟล์ใหม่ที่มีข้อมูลครบเหมือนต้นฉบับได้ แม้จะไม่ได้ใช้
                                 -- คอลัมน์เหล่านี้ในการคำนวณอะไรของระบบเลยก็ตาม (audit/ความครบถ้วน)
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

-- ---------- กลุ่มพื้นที่ Logistic Plan (ตั้งค่าได้ผ่านหน้าเว็บ ไม่ต้องแก้โค้ด+deploy ใหม่) ----------
-- เดิม 4 กลุ่ม (บางบัวทอง/มหาชัย/สุวรรณภูมิ/รอบเช้าต่างจังหวัด) hardcode ไว้ในโค้ด (GROUP_TEMPLATES
-- ใน logistic_plan_export.py) — ย้ายมาเก็บเป็นข้อมูลแทน เพิ่มกลุ่มที่ 5 ได้จากหน้าเว็บโดยตรง
CREATE TABLE IF NOT EXISTS logistic_group (
    id              SERIAL PRIMARY KEY,
    customer_id     INTEGER NOT NULL REFERENCES customer(id),
    group_name      VARCHAR(50) NOT NULL,      -- ชื่อกลุ่มพื้นที่ เช่น "บางบัวทอง"
    template_key    VARCHAR(50) NOT NULL,      -- เชื่อมกับ template_version.template_key เช่น
                                                -- "logistic_บางบัวทอง" — ต้องขึ้นต้นด้วย "logistic_" เสมอ
    sheet_name      VARCHAR(100) NOT NULL,     -- ชื่อ sheet ที่ใช้จริงในไฟล์เทมเพลตของกลุ่มนี้
    display_order   INTEGER NOT NULL DEFAULT 0,
    is_active       BOOLEAN NOT NULL DEFAULT TRUE,  -- ปิดใช้งานกลุ่มได้โดยไม่ต้องลบทิ้ง (กันลบพลาด)
    created_at      TIMESTAMP DEFAULT now(),
    UNIQUE (customer_id, group_name),
    UNIQUE (customer_id, template_key)
);
CREATE INDEX IF NOT EXISTS idx_logistic_group_customer ON logistic_group(customer_id);

-- seed 4 กลุ่มเดิมที่เคย hardcode ไว้ — ทำครั้งเดียว ไม่ทับถ้ามีอยู่แล้ว (สำคัญมาก: ถ้าไม่ seed
-- ตรงนี้ ระบบจะ "ลืม" กลุ่มเดิมทั้งหมดทันทีหลัง migrate เพราะย้ายจาก hardcode มาเป็นข้อมูลแล้ว)
DO $$
DECLARE
    cpall_id INTEGER;
BEGIN
    SELECT id INTO cpall_id FROM customer WHERE code = 'cpall';
    IF cpall_id IS NOT NULL THEN
        INSERT INTO logistic_group (customer_id, group_name, template_key, sheet_name, display_order)
        VALUES
            (cpall_id, 'บางบัวทอง', 'logistic_บางบัวทอง', 'บางบัวทอง-ผลิต', 1),
            (cpall_id, 'มหาชัย', 'logistic_มหาชัย', 'มหาชัย-ผลิต', 2),
            (cpall_id, 'สุวรรณภูมิ', 'logistic_สุวรรณภูมิ', 'สุวรรณภูมิ-ผลิต', 3),
            (cpall_id, 'รอบเช้าต่างจังหวัด', 'logistic_รอบเช้าต่างจังหวัด', 'บางบัวทอง-ผลิต', 4)
        ON CONFLICT (customer_id, group_name) DO NOTHING;
    END IF;
END $$;

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
    original_filename   TEXT,                   -- ชื่อไฟล์ตอนที่ Admin อัปโหลดจริง (เช่น "logistic_plan_บางบัวทอง_v2.xlsx") ไว้ให้ดูย้อนหลังว่าไฟล์นี้คือไฟล์ไหน
    is_active           BOOLEAN NOT NULL DEFAULT FALSE,
    uploaded_at         TIMESTAMP DEFAULT now(),
    validation_summary  TEXT,                   -- เช่น "sku_count=19" เก็บไว้ดูย้อนหลังเฉยๆ
    UNIQUE (customer_id, template_key, version_number)
);
CREATE INDEX IF NOT EXISTS idx_template_version_key ON template_version(customer_id, template_key);

-- ฐานข้อมูลที่เคยรัน schema.sql เวอร์ชันก่อนหน้ามาแล้ว (มี template_version อยู่แล้วแบบไม่มี
-- original_filename) เพิ่มคอลัมน์ให้ — ค่าเก่าที่มีอยู่แล้วจะเป็น NULL (ไม่รู้ชื่อไฟล์ต้นฉบับ เพราะ
-- ตอนอัปโหลดยังไม่ได้เก็บไว้ — ไม่กระทบการทำงานอะไร แค่ช่องว่างเฉยๆ)
ALTER TABLE template_version ADD COLUMN IF NOT EXISTS original_filename TEXT;

-- เชื่อมแผนที่สร้างไว้เข้ากับเทมเพลตเวอร์ชันที่ใช้จริงตอนนั้น — NULL ได้สำหรับแผนเก่าก่อนมีระบบนี้
-- (การ "บันทึกจริง" ว่าใช้เวอร์ชันไหนทำใน sub-phase 3 ตอนเปลี่ยน flow สร้างแผน — คอลัมน์นี้แค่เตรียมที่ไว้)
ALTER TABLE plan_run ADD COLUMN IF NOT EXISTS production_template_version_id INTEGER
    REFERENCES template_version(id);

-- ฐานข้อมูลที่เคยรัน schema.sql เวอร์ชันก่อนหน้ามาแล้ว (มี po_import/po_line อยู่แล้วแบบไม่มีคอลัมน์
-- เก็บข้อมูลครบทุกคอลัมน์ของไฟล์ต้นฉบับ) เพิ่มคอลัมน์ให้ — ค่าเก่าที่มีอยู่แล้วจะเป็น NULL (แผน/PO เก่า
-- ก่อนมีระบบนี้ยังใช้งานได้ปกติ แค่ regenerate ไฟล์แบบครบทุกคอลัมน์ไม่ได้ ต้อง fallback อย่างอื่นแทน)
ALTER TABLE po_import ADD COLUMN IF NOT EXISTS column_order JSONB;
ALTER TABLE po_line ADD COLUMN IF NOT EXISTS all_values JSONB;

-- product_master/location_mapping: เปลี่ยนจาก composite PK (customer_id, barcode/fc_code) เป็น
-- surrogate key (id) ต่างหาก — composite PK ใช้กับ Django Admin ไม่ได้ (ข้อจำกัดที่รู้มาตั้งแต่แรก
-- เพิ่งมาแก้ตอนนี้) ไม่กระทบข้อมูลเดิมเลย (แค่เปลี่ยน PK ไม่ได้ลบ/ย้ายอะไร) ปลอดภัย รันซ้ำได้ —
-- ใช้ชื่อ product_master เพราะ RENAME (sku_master -> product_master) ทำไปแล้วตอนต้นไฟล์เสมอ
DO $$
BEGIN
    IF NOT EXISTS (SELECT 1 FROM information_schema.columns
                    WHERE table_name = 'product_master' AND column_name = 'id') THEN
        ALTER TABLE product_master ADD COLUMN id SERIAL;
        ALTER TABLE product_master DROP CONSTRAINT IF EXISTS sku_master_pkey;
        ALTER TABLE product_master ADD PRIMARY KEY (id);
        ALTER TABLE product_master ADD CONSTRAINT sku_master_customer_barcode_key UNIQUE (customer_id, barcode);
    END IF;

    IF NOT EXISTS (SELECT 1 FROM information_schema.columns
                    WHERE table_name = 'location_mapping' AND column_name = 'id') THEN
        ALTER TABLE location_mapping ADD COLUMN id SERIAL;
        ALTER TABLE location_mapping DROP CONSTRAINT location_mapping_pkey;
        ALTER TABLE location_mapping ADD PRIMARY KEY (id);
        ALTER TABLE location_mapping ADD CONSTRAINT location_mapping_customer_fc_code_key UNIQUE (customer_id, fc_code);
    END IF;
END $$;
ALTER TABLE plan_run_logistic_file ADD COLUMN IF NOT EXISTS template_version_id INTEGER
    REFERENCES template_version(id);

CREATE INDEX IF NOT EXISTS idx_plan_run_import_run ON plan_run_import(plan_run_id);
CREATE INDEX IF NOT EXISTS idx_plan_run_logistic_run ON plan_run_logistic_file(plan_run_id);

-- sku_master/location_mapping: เพิ่ม is_active ให้ปิดใช้งานได้เหมือน logistic_group (สินค้าเลิกขาย/
-- จุดส่งเลิกใช้ ไม่ต้องลบทิ้ง แค่ปิดไว้) ค่าเดิมทั้งหมดเป็น TRUE อัตโนมัติ ไม่กระทบข้อมูลเดิมเลย
ALTER TABLE product_master ADD COLUMN IF NOT EXISTS is_active BOOLEAN NOT NULL DEFAULT TRUE;
ALTER TABLE location_mapping ADD COLUMN IF NOT EXISTS is_active BOOLEAN NOT NULL DEFAULT TRUE;

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
    actual_production_qty NUMERIC(10,2),          -- เฉพาะ production — "ยอดที่ต้องผลิตจริง" (แถวรอง
                                                    -- ใต้ยอดสั่งตาม PO ในเทมเพลต รวมยอดเผื่อ หักลบยอดคืน
                                                    -- แล้ว) ผลจาก LibreOffice จริง — ไม่ใช่แค่ grand_total
    basket_total    INTEGER                      -- เฉพาะ logistic, ผลจาก LibreOffice จริง
);
CREATE INDEX IF NOT EXISTS idx_plan_sku_result_run ON plan_sku_result(plan_run_id);
CREATE INDEX IF NOT EXISTS idx_plan_sku_result_sheet ON plan_sku_result(plan_run_id, sheet_type, group_name);

-- "ยอดที่ต้องผลิตจริง" — column นี้มีอยู่แล้วในนิยาม CREATE TABLE ด้านบน (fresh install ครบอยู่แล้ว)
-- บรรทัดนี้แค่เผื่อฐานข้อมูลเก่าที่เคยมีตารางนี้อยู่ก่อนจะเพิ่มคอลัมน์นี้เข้ามา — ต้องอยู่ "หลัง" CREATE
-- TABLE เสมอ (เจอบั๊กจริง — ตอน fresh install ตารางยังไม่ถูกสร้างเลยตอนบรรทัดนี้รัน ถ้าอยู่ก่อน)
ALTER TABLE plan_sku_result ADD COLUMN IF NOT EXISTS actual_production_qty NUMERIC(10,2);

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

    -- product_master: เพิ่ม customer_id, ใส่ค่าเดิมทั้งหมดเป็นของ cpall, เปลี่ยน PK — ใช้ชื่อ
    -- product_master เพราะ RENAME (sku_master -> product_master) ทำไปเป็นขั้นแรกสุดแล้วเสมอ
    IF EXISTS (SELECT 1 FROM information_schema.tables WHERE table_name = 'product_master') THEN
        IF NOT EXISTS (SELECT 1 FROM information_schema.columns
                        WHERE table_name = 'product_master' AND column_name = 'customer_id') THEN
            ALTER TABLE product_master ADD COLUMN customer_id INTEGER REFERENCES customer(id);
            UPDATE product_master SET customer_id = cpall_id WHERE customer_id IS NULL;
            ALTER TABLE product_master ALTER COLUMN customer_id SET NOT NULL;
            ALTER TABLE product_master DROP CONSTRAINT IF EXISTS sku_master_pkey;
            ALTER TABLE product_master ADD PRIMARY KEY (customer_id, barcode);
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

ALTER TABLE product_master ENABLE ROW LEVEL SECURITY;
ALTER TABLE product_master FORCE ROW LEVEL SECURITY;
DROP POLICY IF EXISTS customer_isolation ON product_master;
CREATE POLICY customer_isolation ON product_master
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

ALTER TABLE logistic_group ENABLE ROW LEVEL SECURITY;
ALTER TABLE logistic_group FORCE ROW LEVEL SECURITY;
DROP POLICY IF EXISTS customer_isolation ON logistic_group;
CREATE POLICY customer_isolation ON logistic_group
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
