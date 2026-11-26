import contextlib
import io
import tempfile
import unittest
from datetime import date
from pathlib import Path

from openpyxl import Workbook

from nohin.cli import main
from nohin.config import ReviewTableConfig
from nohin.review_table import (ROW_NO_FINDING_MARK, ROW_NOT_REQUIRED, ROW_REQUIRED, ROW_STRAY, list_drawings,
                                range_label, read_review_table)
from tests.dummy import COMPLETE, write_review_record

CFG = ReviewTableConfig()

try:
    import PIL  # noqa: F401
    from openpyxl.drawing.image import Image as XLImage
    HAS_PIL = True
except ImportError:  # Pillow がない環境では画像入りのテストを飛ばす
    HAS_PIL = False


class ReviewTableTest(unittest.TestCase):
    def setUp(self):
        self.tmp = tempfile.TemporaryDirectory()
        self.dir = Path(self.tmp.name)

    def tearDown(self):
        self.tmp.cleanup()

    def make(self, *args, **kw) -> Path:
        return write_review_record(self.dir / "rv.xlsx", *args, **kw)

    def test_complete(self):
        res = read_review_table(self.make(), CFG)
        self.assertTrue(res.readable and res.has_table)
        self.assertEqual([r.kind for r in res.rows], [ROW_REQUIRED, ROW_REQUIRED])  # 「-」だけの行と備考は数えない
        self.assertTrue(res.action_done)
        self.assertEqual(res.latest_date, date(2026, 11, 20))
        self.assertEqual(len(res.approvals), 1)
        self.assertEqual(range_label(res.approvals[0][1]), "H3:H4")
        self.assertEqual(res.columns["date"], 7)  # 処置欄より右の「日付」（左の指摘日ではない）

    def test_header_follows_row_shift_and_merge(self):
        res = read_review_table(self.make(extra_upper_rows=5, header_merged=True), CFG)
        base = read_review_table(write_review_record(self.dir / "b.xlsx"), CFG)
        self.assertEqual(res.header_row, base.header_row + 5)
        self.assertEqual(res.data_start, res.header_row + 2)
        self.assertEqual(len(res.findings), 2)

    def test_row_kinds(self):
        rows = [("指摘なし", "", None, ""), ("A", "直した", date(2026, 11, 1), ""), ("B", "", None, "－"),
                ("C", "回答", None, "-"), ("", "処置だけ", None, ""), ("D", "直した", "未定", "")]
        res = read_review_table(self.make(rows), CFG)
        self.assertEqual([r.kind for r in res.rows],
                         [ROW_NO_FINDING_MARK, ROW_REQUIRED, ROW_NOT_REQUIRED, ROW_NOT_REQUIRED, ROW_STRAY, ROW_REQUIRED])
        self.assertEqual([r.no for r in res.missing_dates()], ["6"])
        self.assertFalse(res.action_done)
        self.assertEqual(len(res.strays), 1)

    def test_not_required_rows_do_not_block(self):
        res = read_review_table(self.make([COMPLETE[0], ("質問", "", None, "－")]), CFG)
        self.assertTrue(res.action_done)
        self.assertEqual(len(res.findings), 2)

    def test_no_table_and_no_approval(self):
        self.assertFalse(read_review_table(self.make(with_table=False), CFG).has_table)
        self.assertEqual(read_review_table(self.make(with_approval=False), CFG).approvals, [])

    def test_unreadable(self):
        bad = self.dir / "bad.xlsx"
        bad.write_bytes(b"not a zip")
        self.assertFalse(read_review_table(bad, CFG).readable)
        self.assertFalse(read_review_table(self.dir / "old.xls", CFG).readable)

    def test_custom_labels(self):
        wb = Workbook()
        ws = wb.active
        ws["B2"] = "承認"
        for c, t in enumerate(["指摘内容", "処置内容", "処置日", "自社印"], start=1):
            ws.cell(5, c, t)
        ws.cell(6, 1, "x")
        ws.cell(6, 2, "y")
        ws.cell(6, 3, "2026/11/01")
        p = self.dir / "c.xlsx"
        wb.save(p)
        res = read_review_table(p, CFG)
        self.assertEqual(res.header_row, 5)
        self.assertEqual(res.latest_date, date(2026, 11, 1))

    def test_stops_on_huge_sheet(self):
        # 表の下の遠い行に値があっても、空行が続いたら打ち切る（最終行まで読まない）
        p = self.make()
        from openpyxl import load_workbook
        wb = load_workbook(p)
        wb.active.cell(100000, 3, "遠い行")
        wb.save(p)
        res = read_review_table(p, CFG)
        self.assertEqual(len(res.findings), 2)

    @unittest.skipUnless(HAS_PIL, "Pillow がないため画像のテストを省略")
    def test_list_drawings(self):
        from PIL import Image
        png = self.dir / "s.png"
        Image.new("RGB", (20, 20), "red").save(png)
        p = self.make()
        from openpyxl import load_workbook
        wb = load_workbook(p)
        wb.active.add_image(XLImage(str(png)), "H14")
        wb.save(p)
        objs = list_drawings(p)
        self.assertEqual(len(objs), 1)
        self.assertEqual((objs[0].kind, objs[0].from_cell), ("画像", (14, 8)))

    def test_inspect_command(self):
        p = self.make([COMPLETE[0], ("B", "", None, "－")])
        out = io.StringIO()
        with contextlib.redirect_stdout(out):
            code = main(["inspect", str(p)])
        self.assertEqual(code, 0)
        text = out.getvalue()
        for s in ["見出し", "承認欄: ラベル H2", "指摘（捺印不要）", "最新の処置日（社内の承認印の日付）: 2026-11-18", "画像・図形: 0 件"]:
            self.assertIn(s, text)


if __name__ == "__main__":
    unittest.main()
