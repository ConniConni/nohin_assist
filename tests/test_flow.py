"""collect → 手で更新 → report の一連の流れと、管理表の扱いを確認する。"""

import contextlib
import io
import tempfile
import unittest
from pathlib import Path
from unittest import mock

from openpyxl import load_workbook

from nohin import progress
from nohin.cli import main
from nohin.config import ConfigError, load_config, load_delivery
from nohin.manifest import ManifestError, ensure_not_open, read_manifest
from nohin.model import TARGET_CHECK, TARGET_EXCLUDE
from nohin.scan import scan_project
from tests.dummy import make_clean_project, make_demo, touch, write_review_record


def run(*args) -> tuple[int, str]:
    out, err = io.StringIO(), io.StringIO()
    with contextlib.redirect_stdout(out), contextlib.redirect_stderr(err):
        code = main(list(args))
    return code, out.getvalue() + err.getvalue()


class FlowTest(unittest.TestCase):
    def setUp(self):
        self.tmp = tempfile.TemporaryDirectory()
        self.demo = make_demo(Path(self.tmp.name) / "demo")
        self.dfile = str(self.demo / "delivery.toml")
        self.cfg = load_config(self.demo / "config.toml")
        self.delivery = load_delivery(self.dfile, self.cfg)
        self.mpath = self.delivery.manifest_path

    def tearDown(self):
        self.tmp.cleanup()

    def collect(self, *extra):
        code, out = run("collect", "-d", self.dfile, *extra)
        self.assertEqual(code, 0, out)
        return out

    def issues(self):
        wb = load_workbook(self.mpath)
        return [r for r in wb["警告"].iter_rows(min_row=2, values_only=True)]

    def test_scan_excludes_back_and_temp(self):
        files, _ = scan_project(self.delivery.projects[0], self.cfg)
        rels = [f.rel for f in files]
        self.assertFalse(any("/back/" in r for r in rels))
        self.assertFalse(any("~$" in r for r in rels))
        files_b, _ = scan_project(self.delivery.projects[1], self.cfg)
        self.assertTrue(any("/20261101/" in f.rel for f in files_b))  # パターン2も拾う

    def test_collect_only_deliverable_locations(self):
        # 工程フォルダ直下と 成果物/内部Rv・外部Rv（日付フォルダ）だけを載せる。レビューフォルダのドキュメントは載せない
        self.collect()
        records, _ = read_manifest(self.mpath)
        rels = {r.rel for r in records if r.project == "案件A"}
        self.assertNotIn("040.設計/作業用/機能設計書_A_作業中.xlsx", rels)
        self.assertNotIn("040.設計/成果物/外部Rv/機能設計書_A.xlsx", rels)
        self.assertFalse(any("/過去分/" in r for r in rels))
        self.assertIn("040.設計/機能設計書_A.xlsx", rels)
        self.assertIn("040.設計/成果物/外部Rv/レビュー記録表(社外)(1回目)_設計_A.xlsx", rels)
        code, out = run("collect", "-d", self.dfile, "--check-names")
        self.assertNotIn("作業用", out)
        self.assertNotIn("成果物/外部Rv/機能設計書_A.xlsx", out)

    def test_allow_multiple_per_project(self):
        # 案件A は allow_multiple で機能設計書が複数あってよい。指定を外すと重複として出る
        self.collect()
        dup = [r for r in self.issues() if r[2] == "ドキュメント重複"]
        self.assertEqual([(r[1], "結合試験仕様書" in r[3]) for r in dup], [("案件C", True)])

        Path(self.dfile).write_text(Path(self.dfile).read_text(encoding="utf-8")
                                    .replace('allow_multiple = ["機能設計書"]\n', ""), encoding="utf-8")
        self.collect("--new")
        dup = {(r[1], r[3].split("に")[1].split("が")[0]) for r in self.issues() if r[2] == "ドキュメント重複"}
        self.assertEqual(dup, {("案件A", "機能設計書"), ("案件C", "結合試験仕様書")})

    def test_collect_detects_expected_issues(self):
        self.collect()
        cats = {(r[1], r[2]) for r in self.issues()}
        for expected in [("案件A", "社内外不一致"), ("案件A", "ペア不足"), ("案件A", "規則外"),
                         ("案件A", "全角括弧"), ("案件B", "回数の抜け"), ("案件C", "ドキュメント重複"),
                         ("案件C", "ドキュメントなし")]:
            self.assertIn(expected, cats)
        # 対象工程を絞った案件で、対象外工程の行が出ないこと
        records, _ = read_manifest(self.mpath)
        self.assertEqual({r.process for r in records if r.project == "案件B"}, {"030", "040"})

    def test_check_names_does_not_write(self):
        code, out = run("collect", "-d", self.dfile, "--check-names")
        self.assertEqual(code, 0)
        self.assertIn("判定結果", out)
        self.assertFalse(self.mpath.exists())

    def test_merge_keeps_human_edits_and_confirmations(self):
        self.collect()
        wb = load_workbook(self.mpath)
        ws = wb["管理表"]
        hm = {c.value: c.column for c in ws[1]}
        ws.cell(2, hm["自社捺印"], "済")
        ws.cell(2, hm["備考"], "手で記入")
        wi = wb["警告"]
        wi.cell(2, 6, "確認済")
        confirmed_msg = wi.cell(2, 4).value
        wb.save(self.mpath)

        touch(self.demo / "src" / "案件A" / "050.製造" / "成果物" / "内部Rv" / "レビュー記録表(社内)(2回目)_CD_A.xlsx")
        out = self.collect()
        self.assertIn("新規 1 件", out)
        records, confirmed = read_manifest(self.mpath)
        first = records[0]
        self.assertEqual((first.st_self, first.note), ("済", "手で記入"))
        self.assertTrue(any("(2回目)_CD_A" in r.filename for r in records))
        self.assertIn("確認済", confirmed.values())
        self.assertTrue(any(r[3] == confirmed_msg and r[5] == "確認済" for r in self.issues()))
        self.assertTrue(list((self.delivery.work_dir / "_backup").iterdir()))

    def test_missing_source_reported(self):
        self.collect()
        (self.demo / "src" / "案件A" / "030.調査" / "調査検討書_A.xlsx").unlink()
        run("report", "-d", self.dfile)
        self.assertIn("元ファイルなし", {r[2] for r in self.issues()})

    def test_report_completion(self):
        # 規則どおりの1案件だけにする
        root = make_clean_project(Path(self.tmp.name) / "clean" / "案件X", codes=("040", "060"))
        dfile = Path(self.tmp.name) / "clean" / "delivery.toml"
        dfile.write_text(f'[delivery]\nname = "X"\nwork_dir = "work"\n[[projects]]\nname = "案件X"\nroot = "{root.as_posix()}"\n',
                         encoding="utf-8")
        (dfile.parent / "config.toml").write_text((self.demo / "config.toml").read_text(encoding="utf-8"), encoding="utf-8")
        code, out = run("collect", "-d", str(dfile))
        self.assertIn("判定: 未完了", out)
        mpath = load_delivery(dfile).manifest_path

        wb = load_workbook(mpath)
        ws = wb["管理表"]
        hm = {c.value: c.column for c in ws[1]}
        for row in range(2, ws.max_row + 1):
            for h in ("指摘展開", "自社捺印", "取引先送付", "取引先捺印", "納品コピー"):
                if ws.cell(row, hm[h]).value == "未":
                    ws.cell(row, hm[h], "済")
        wb.save(mpath)
        code, out = run("report", "-d", str(dfile))
        self.assertIn("判定: 完了", out, out)
        self.assertEqual(load_workbook(mpath)["サマリー"]["B4"].value, "完了")

    def test_order_warning_and_invalid_value(self):
        self.collect()
        wb = load_workbook(self.mpath)
        ws = wb["管理表"]
        hm = {c.value: c.column for c in ws[1]}
        # 機能設計書（自社・取引先とも必要）の行で、自社捺印前に納品コピーを済にする
        row = next(r for r in range(2, ws.max_row + 1) if ws.cell(r, hm["ファイル名"]).value == "機能設計書_A.xlsx")
        ws.cell(row, hm["納品コピー"], "済")
        ws.cell(row + 1, hm["自社捺印"], "OK")
        wb.save(self.mpath)
        run("report", "-d", self.dfile)
        cats = {r[2] for r in self.issues()}
        self.assertIn("工程順序", cats)
        self.assertIn("入力値", cats)

    def test_report_rereads_review_records(self):
        # 記録表の処置を埋めて report を実行すると、状況・進捗が更新され、手で書いた備考は残る
        self.collect()
        name = "レビュー記録表(社外)(1回目)_設計_A.xlsx"
        wb = load_workbook(self.mpath)
        ws = wb["管理表"]
        hm = {c.value: c.column for c in ws[1]}
        row = next(r for r in range(2, ws.max_row + 1) if ws.cell(r, hm["ファイル名"]).value == name)
        self.assertEqual(ws.cell(row, hm["記録表の状況"]).value, "処置中")
        ws.cell(row, hm["備考"], "取引先に確認中")
        wb.save(self.mpath)

        write_review_record(self.demo / "src" / "案件A" / "040.設計" / "成果物" / "外部Rv" / name)  # 処置完了の内容に
        run("report", "-d", self.dfile)
        rec = next(r for r in read_manifest(self.mpath)[0] if r.filename == name)
        self.assertEqual((rec.rv_state, rec.rv_progress, rec.note), ("処置完了", "指摘2・処置2・日付2", "取引先に確認中"))

    def test_one_broken_record_does_not_stop(self):
        # 記録表の読み取りで想定外のエラーが起きても、その行を「読取不可」にして処理を続ける
        target = "レビュー記録表(社内)(1回目)_CD_A.xlsx"
        original = progress.read_review_table

        def flaky(path, cfg):
            if Path(path).name == target:
                raise RuntimeError("想定外")
            return original(path, cfg)

        with mock.patch.object(progress, "read_review_table", flaky):
            self.collect()
        rec = next(r for r in read_manifest(self.mpath)[0] if r.filename == target)
        self.assertEqual(rec.rv_state, "読取不可")
        self.assertTrue(any(r[2] == "記録表読取不可" and "RuntimeError" in r[3] for r in self.issues()))

    def test_share_order_exception(self):
        # 社外レビュー記録表: 指摘展開が未でも自社捺印の済は警告しない。取引先送付の済は警告する
        self.collect()
        wb = load_workbook(self.mpath)
        ws = wb["管理表"]
        hm = {c.value: c.column for c in ws[1]}
        name = "レビュー記録表(社外)(1回目)_設計_B.xlsx"
        row = next(r for r in range(2, ws.max_row + 1) if ws.cell(r, hm["ファイル名"]).value == name)
        no = ws.cell(row, hm["No"]).value
        ws.cell(row, hm["自社捺印"], "済")
        wb.save(self.mpath)
        run("report", "-d", self.dfile)
        order = [r for r in self.issues() if r[2] == "工程順序" and str(no) in (r[4] or "")]
        self.assertEqual(order, [])

        wb = load_workbook(self.mpath)
        wb["管理表"].cell(row, hm["取引先送付"], "済")
        wb.save(self.mpath)
        run("report", "-d", self.dfile)
        order = [r for r in self.issues() if r[2] == "工程順序" and str(no) in (r[4] or "")]
        self.assertEqual(len(order), 1)
        self.assertIn("指摘展開", order[0][3])

    def test_locked_manifest(self):
        self.collect()
        lock = self.mpath.parent / f"~${self.mpath.name}"
        lock.write_text("")
        with self.assertRaises(ManifestError):
            ensure_not_open(self.mpath)
        code, out = run("report", "-d", self.dfile)
        self.assertEqual(code, 2)
        self.assertIn("開かれています", out)

    def test_broken_header(self):
        self.collect()
        wb = load_workbook(self.mpath)
        hm = {c.value: c.column for c in wb["管理表"][1]}
        wb["管理表"].cell(1, hm["ID"], "IDX")
        wb.save(self.mpath)
        code, out = run("report", "-d", self.dfile)
        self.assertEqual(code, 2)
        self.assertIn("見出しが足りません", out)

    def test_needs_check_rows(self):
        self.collect()
        records, _ = read_manifest(self.mpath)
        checks = {r.filename for r in records if r.target == TARGET_CHECK}
        excluded = {r.filename for r in records if r.target == TARGET_EXCLUDE}
        self.assertIn("メモ.txt", excluded)
        self.assertIn("レビュー記録表_設計_A.xlsx", checks)
        self.assertIn("レビュー記録表(社内)(1回目)_UD作成_A2.xlsx", checks)


class ConfigErrorTest(unittest.TestCase):
    def test_bad_regex_and_missing(self):
        with tempfile.TemporaryDirectory() as d:
            p = Path(d) / "delivery.toml"
            p.write_text('[delivery]\nname="x"\n', encoding="utf-8")
            code, out = run("collect", "-d", str(p), "-c", str(Path(d) / "none.toml"))
            self.assertEqual(code, 2)
            self.assertIn("見つかりません", out)

    def test_allow_multiple_unknown_name(self):
        cfg = load_config(Path(__file__).resolve().parent.parent / "config.example.toml")
        with tempfile.TemporaryDirectory() as d:
            p = Path(d) / "delivery.toml"
            p.write_text('[delivery]\nname="x"\nwork_dir="w"\n[[projects]]\nname="A"\nroot="a"\n'
                         'allow_multiple = ["基本設計書"]\n', encoding="utf-8")
            with self.assertRaisesRegex(ConfigError, "allow_multiple"):
                load_delivery(p, cfg)
            p.write_text('[delivery]\nname="x"\nwork_dir="w"\n[[projects]]\nname="A"\nroot="a"\n'
                         'allow_multiple = "機能設計書"\n', encoding="utf-8")
            with self.assertRaisesRegex(ConfigError, "配列"):
                load_delivery(p, cfg)


if __name__ == "__main__":
    unittest.main()
