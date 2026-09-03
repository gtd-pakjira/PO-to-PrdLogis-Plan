"""
test_plan_view_data.py — เทสสูตรคำนวณที่จำลองมาจากสูตร Excel จริง (ยืนยันกับไฟล์จริงไปแล้วตอนพัฒนา
ครั้งแรก — เทสพวกนี้ล็อกพฤติกรรมไว้ กันคนแก้ทีหลังพลาดโดยไม่รู้ตัว) ไม่แตะ database เลย
"""
from django.test import SimpleTestCase

from customers.cpall.logic.plan_view_data import _basket_total, _pack_breakdown_text


class PackBreakdownTextTests(SimpleTestCase):
    """จำลองสูตร Excel: IF(qty=0,"",IF(MOD(qty,pack)=0, qty/pack, IF(INT(qty/pack)=0, MOD&" P", INT&" + "&MOD&" P")))"""

    def test_exact_multiple_of_pack_size(self):
        self.assertEqual(_pack_breakdown_text(72, 36), "2")

    def test_whole_packs_plus_remainder(self):
        # ค่าจริงที่เคยเทียบกับไฟล์ Admin: 133 ชิ้น บรรจุ 36/ลัง -> "3 + 25 P"
        self.assertEqual(_pack_breakdown_text(133, 36), "3 + 25 P")

    def test_less_than_one_full_pack(self):
        self.assertEqual(_pack_breakdown_text(25, 36), "25 P")

    def test_zero_qty_returns_empty_string(self):
        self.assertEqual(_pack_breakdown_text(0, 36), "")

    def test_none_qty_returns_empty_string(self):
        self.assertEqual(_pack_breakdown_text(None, 36), "")

    def test_zero_pack_size_returns_empty_string(self):
        self.assertEqual(_pack_breakdown_text(50, 0), "")


class BasketTotalTests(SimpleTestCase):
    """จำลองสูตร Excel: =SUMPRODUCT(ROUNDUP(qty/pack, 0)) — ปัดเศษขึ้นทีละคอลัมน์แล้วรวม"""

    def test_sums_ceiling_per_column(self):
        # ceil(133/36)=4, ceil(102/36)=3, ceil(56/36)=2 -> รวม 9
        qty_by_column = {"บางบัวทอง PO1": 133, "ชลบุรี PO1": 102, "ชลบุรี PO2": 56}
        self.assertEqual(_basket_total(qty_by_column, 36), 9)

    def test_zero_pack_size_returns_zero(self):
        self.assertEqual(_basket_total({"A": 10}, 0), 0)

    def test_empty_columns_returns_zero(self):
        self.assertEqual(_basket_total({}, 36), 0)

    def test_skips_falsy_values_without_error(self):
        # คอลัมน์ที่ไม่มียอดสั่ง (0 หรือ None) ไม่ควรถูกนับหรือทำให้ error
        self.assertEqual(_basket_total({"A": 0, "B": None, "C": 36}, 36), 1)
