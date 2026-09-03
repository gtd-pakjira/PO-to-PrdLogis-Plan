"""
test_date_utils.py — เทสฟังก์ชันแปลง/จัดรูปแบบวันที่ (ไม่แตะ database เลย — SimpleTestCase บังคับ
ไม่ให้แตะ DB ด้วย กันเผลอเขียนเทสที่พึ่ง DB โดยไม่รู้ตัว)
"""
from datetime import date

from django.test import SimpleTestCase

from customers.cpall.logic.date_utils import format_thai_short, parse_date_arg


class ParseDateArgTests(SimpleTestCase):
    def test_iso_format(self):
        self.assertEqual(parse_date_arg("2026-08-28"), date(2026, 8, 28))

    def test_thai_slash_format(self):
        self.assertEqual(parse_date_arg("28/08/2026"), date(2026, 8, 28))

    def test_dash_format(self):
        self.assertEqual(parse_date_arg("28-08-2026"), date(2026, 8, 28))

    def test_strips_whitespace(self):
        self.assertEqual(parse_date_arg("  2026-08-28  "), date(2026, 8, 28))

    def test_invalid_format_raises(self):
        with self.assertRaises(ValueError):
            parse_date_arg("ไม่ใช่วันที่")


class FormatThaiShortTests(SimpleTestCase):
    def test_converts_to_buddhist_era_two_digit(self):
        # 2026 ค.ศ. = 2569 พ.ศ. -> 2 หลักท้าย = 69
        self.assertEqual(format_thai_short(date(2026, 8, 28)), "28/08/69")

    def test_pads_single_digit_day_and_month(self):
        self.assertEqual(format_thai_short(date(2026, 1, 5)), "05/01/69")

    def test_year_rollover(self):
        # 2000 ค.ศ. = 2543 พ.ศ. -> 43
        self.assertEqual(format_thai_short(date(2000, 12, 31)), "31/12/43")
