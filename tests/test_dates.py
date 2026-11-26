import unittest
from datetime import date, datetime

from nohin.dates import parse_date


class ParseDateTest(unittest.TestCase):
    def test_readable(self):
        d = date(2026, 11, 20)
        for v in [d, datetime(2026, 11, 20, 9, 30), "2026/11/20", "2026-11-20", "2026.11.20", "2026年11月20日",
                  "２０２６／１１／２０", " 2026/11/20 ", "2026/11/20 10:00", "令和8年11月20日", "R8.11.20", "r8/11/20"]:
            self.assertEqual(parse_date(v), d, repr(v))
        self.assertEqual(parse_date("令和元年5月1日"), date(2019, 5, 1))
        self.assertEqual(parse_date("H31.4.30"), date(2019, 4, 30))

    def test_unreadable(self):
        for v in [None, "", "11/20", "2026/13/01", "2026/2/30", "未定", 46000, 3.5, "済"]:
            self.assertIsNone(parse_date(v), repr(v))


if __name__ == "__main__":
    unittest.main()
