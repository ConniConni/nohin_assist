"""デモデータを正誤表と突き合わせ、全項目が一致することを確認する。"""

import contextlib
import io
import tempfile
import unittest
from pathlib import Path

from nohin.cli import main as nohin_main
from tests.dummy import make_demo
from tests.verify_demo import to_markdown, verify, verify_deliver

ROOT = Path(__file__).resolve().parent.parent


class DemoVerifyTest(unittest.TestCase):
    def setUp(self):
        self.tmp = tempfile.TemporaryDirectory()
        self.demo = make_demo(Path(self.tmp.name) / "demo")
        with contextlib.redirect_stdout(io.StringIO()):
            nohin_main(["collect", "-d", str(self.demo / "delivery.toml")])

    def tearDown(self):
        self.tmp.cleanup()

    def test_all_cases_match(self):
        ng = [r for r in verify(self.demo) if not r.ok]
        self.assertEqual(ng, [], "\n".join(f"{r.id}: 期待={r.expected} / 実際={r.actual}" for r in ng))

    def test_deliver_stage_matches(self):
        # 試行手順と同じ順: check（止まる）→ overwrite
        dfile = str(self.demo / "delivery.toml")
        with contextlib.redirect_stdout(io.StringIO()):
            self.assertEqual(nohin_main(["deliver", "-d", dfile, "--mode", "check"]), 1)
            self.assertEqual(nohin_main(["deliver", "-d", dfile, "--mode", "overwrite"]), 0)
        ng = [r for r in verify_deliver(self.demo) if not r.ok]
        self.assertEqual(ng, [], "\n".join(f"{r.id}: 期待={r.expected} / 実際={r.actual}" for r in ng))

    def test_deliver_stage_detects_mismatch(self):
        # deliver を実行しないまま照合すると × になること
        ng = {r.id for r in verify_deliver(self.demo) if not r.ok}
        self.assertTrue({"D-1", "D-2", "T-2"} <= ng)

    def test_detects_mismatch(self):
        # 仕込みを1つ消すと、照合で × になること（照合が素通りしていないかの確認）
        (self.demo / "src" / "案件A" / "040.設計" / "メモ.txt").unlink()
        with contextlib.redirect_stdout(io.StringIO()):
            nohin_main(["collect", "-d", str(self.demo / "delivery.toml"), "--new"])
        ng = {r.id for r in verify(self.demo) if not r.ok}
        self.assertIn("A-4", ng)
        self.assertIn("T-1", ng)

    def test_markdown_is_up_to_date(self):
        doc = ROOT / "docs" / "デモ正誤表.md"
        self.assertEqual(doc.read_text(encoding="utf-8"), to_markdown(),
                         "docs/デモ正誤表.md が古い: python -m tests.verify_demo demo --markdown docs/デモ正誤表.md で再生成する")


if __name__ == "__main__":
    unittest.main()
