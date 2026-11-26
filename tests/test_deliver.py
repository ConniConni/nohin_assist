"""④ 納品コピー（deliver）・restore・納品先のチェックのテスト。demo のダミーを使う。"""

import contextlib
import io
import tempfile
import unicodedata
import unittest
from pathlib import Path
from unittest import mock

from openpyxl import load_workbook

from nohin import fileops
from nohin.cli import main
from nohin.config import load_config, load_delivery
from nohin.fileops import norm_key
from nohin.manifest import read_manifest
from tests.dummy import cl_name, make_demo, touch

CL_DIFF = f"案件A/040.設計/成果物/{cl_name('040', '社外', 1, 'A')}"   # demo で納品先に違う内容を置いた行
DOC_SAME = "案件A/030.調査/調査検討書_A.xlsx"                          # demo で納品先に同じ内容を置いた行
CL_NEW = f"案件B/030.調査/成果物/{cl_name('030', '社内', 1, 'B')}"     # 納品先にまだない行
MISNAMED = "レビュー記録表_設計_A.xlsx"                                # 規則外（要確認）の行


def run(*args) -> tuple[int, str]:
    out, err = io.StringIO(), io.StringIO()
    with contextlib.redirect_stdout(out), contextlib.redirect_stderr(err):
        code = main(list(args))
    return code, out.getvalue() + err.getvalue()


class DeliverTest(unittest.TestCase):
    def setUp(self):
        self.tmp = tempfile.TemporaryDirectory()
        self.demo = make_demo(Path(self.tmp.name) / "demo")
        self.dfile = str(self.demo / "delivery.toml")
        self.cfg = load_config(self.demo / "config.toml")
        self.delivery = load_delivery(self.dfile, self.cfg)
        self.mpath = self.delivery.manifest_path
        self.dest = self.demo / "dest"
        self.work = self.delivery.work_dir
        code, out = run("collect", "-d", self.dfile)
        self.assertEqual(code, 0, out)

    def tearDown(self):
        self.tmp.cleanup()

    # ---- 補助
    def deliver(self, mode="list"):
        return run("deliver", "-d", self.dfile, "--mode", mode)

    def rec(self, dest_rel=None, filename=None):
        for r in read_manifest(self.mpath)[0]:
            if (dest_rel and r.dest_rel == dest_rel) or (filename and r.filename == filename):
                return r
        raise AssertionError(f"行がありません: {dest_rel or filename}")

    def edit(self, filename, **values):
        """管理表の行を人が直したことにする（見出し → 値）。"""
        wb = load_workbook(self.mpath)
        ws = wb["管理表"]
        hm = {c.value: c.column for c in ws[1]}
        row = next(r for r in range(2, ws.max_row + 1) if ws.cell(r, hm["ファイル名"]).value == filename)
        for header, v in values.items():
            ws.cell(row, hm[header], v)
        wb.save(self.mpath)

    def issues(self):
        ws = load_workbook(self.mpath)["警告"]
        return [tuple("" if v is None else str(v) for v in row[:5]) for row in ws.iter_rows(min_row=2, values_only=True)]

    def cats(self):
        return [i[2] for i in self.issues()]

    def snapshot(self):
        files = [p for base in (self.demo / "src", self.dest, self.work) for p in base.rglob("*") if p.is_file()]
        return {str(p): p.stat().st_mtime_ns for p in files}

    # ---- list
    def test_list_changes_nothing(self):
        before = self.snapshot()
        code, out = self.deliver("list")
        self.assertEqual(code, 0, out)
        self.assertIn("新規 23 件 / 同じ内容 1 件 / 内容が違う 1 件", out)
        self.assertIn("何も変更していません", out)
        self.assertEqual(self.snapshot(), before)
        self.assertFalse((self.work / "_log").exists())

    # ---- check
    def test_check_stops_when_any_content_differs(self):
        code, out = self.deliver("check")
        self.assertEqual(code, 1, out)
        self.assertIn("1件もコピーせずに止めました", out)
        self.assertFalse((self.dest / CL_NEW).exists())
        self.assertEqual(self.rec(DOC_SAME).st_copy, "未")
        self.assertTrue(list((self.work / "_log").glob("deliver-check_*.txt")))

    def test_check_copies_and_marks_same_content(self):
        (self.dest / CL_DIFF).unlink()
        code, out = self.deliver("check")
        self.assertEqual(code, 0, out)
        new, same = self.rec(CL_NEW), self.rec(DOC_SAME)
        self.assertEqual((new.st_copy, same.st_copy), ("済", "済"))
        self.assertEqual(new.copy_hash, fileops.sha256_of(self.dest / CL_NEW))
        self.assertTrue(new.copy_at)
        self.assertEqual(same.backup, "")  # 同じ内容は上書きしないので退避もしない
        self.assertFalse(list(self.dest.rglob("*" + fileops.TMP_SUFFIX)))
        # 前の工程が未の行はコピーされない
        self.assertFalse((self.dest / "案件A" / "040.設計" / "機能設計書_A.xlsx").exists())
        self.assertEqual(self.rec(filename="機能設計書_A.xlsx").st_copy, "未")

    # ---- overwrite
    def test_overwrite_backs_up_then_replaces(self):
        code, out = self.deliver("overwrite")
        self.assertEqual(code, 0, out)
        r = self.rec(CL_DIFF)
        self.assertEqual(r.st_copy, "済")
        backup = Path(r.backup)
        self.assertEqual(backup.read_bytes(), b"old version")
        self.assertIn("_納品コピー", backup.parent.parent.parent.parent.name)
        self.assertEqual(backup.relative_to(self.work / "_backup").parts[1:], tuple(CL_DIFF.split("/")))
        self.assertEqual((self.dest / CL_DIFF).read_bytes(), Path(r.src).read_bytes())
        self.assertTrue(list(backup.parents[3].glob("納品管理表_*.xlsx")))  # 書き込む前の管理表も退避
        self.assertTrue(list((self.work / "_log").glob("deliver-overwrite_*.txt")))
        # 原本は変わらない
        self.assertEqual(Path(r.src).read_bytes(), b"dummy")

    def test_hash_mismatch_marks_ng_and_keeps_destination(self):
        real_copy = fileops.shutil.copyfile

        def broken(src, dst, *a, **k):
            real_copy(src, dst)
            Path(dst).write_bytes(b"broken")

        with mock.patch.object(fileops.shutil, "copyfile", broken):
            code, out = self.deliver("overwrite")
        self.assertEqual(code, 1, out)
        self.assertIn("一致しません", out)
        self.assertEqual(self.rec(CL_NEW).st_copy, "NG")
        self.assertFalse((self.dest / CL_NEW).exists())
        self.assertEqual((self.dest / CL_DIFF).read_bytes(), b"old version")  # 置き換える前に失敗したので元のまま
        self.assertFalse(list(self.dest.rglob("*" + fileops.TMP_SUFFIX)))
        self.assertIn("NG", self.cats())

    def test_open_manifest_stops_before_copy(self):
        (self.mpath.parent / f"~${self.mpath.name}").write_text("")
        code, out = self.deliver("overwrite")
        self.assertEqual(code, 2)
        self.assertIn("開かれています", out)
        self.assertFalse((self.dest / CL_NEW).exists())
        self.assertEqual((self.dest / CL_DIFF).read_bytes(), b"old version")

    def test_source_being_edited_is_skipped(self):
        src = Path(self.rec(CL_NEW).src)
        touch(src.parent / f"~${src.name}")
        (self.dest / CL_DIFF).unlink()
        code, out = self.deliver("check")
        self.assertIn("元ファイルが Excel で開かれている", out)
        self.assertFalse((self.dest / CL_NEW).exists())

    # ---- restore
    def test_restore_returns_backup_and_resets_status(self):
        self.deliver("overwrite")
        no = self.rec(CL_DIFF).no
        code, out = run("restore", "-d", self.dfile, str(no))
        self.assertEqual(code, 0, out)
        self.assertEqual((self.dest / CL_DIFF).read_bytes(), b"old version")
        r = self.rec(CL_DIFF)
        self.assertEqual((r.st_copy, r.copy_at, r.copy_hash), ("未", "", ""))
        self.assertIn("_restore", r.backup)
        self.assertEqual(Path(r.backup).read_bytes(), b"dummy")  # 戻す前の納品先も退避されている
        self.assertTrue(list((self.work / "_log").glob("restore_*.txt")))

    def test_restore_refuses_other_backups_and_new_rows(self):
        self.deliver("overwrite")
        new_no = self.rec(CL_NEW).no
        code, out = run("restore", "-d", self.dfile, str(new_no))
        self.assertEqual(code, 1)
        self.assertIn("バックアップがありません", out)
        # 管理表のバックアップ（_collect）など、納品コピー以外のバックアップは戻さない
        other = next((self.work / "_backup").glob("*_collect*"), None) or (self.work / "_backup" / "x_collect")
        self.edit(cl_name("040", "社外", 1, "A"), **{"バックアップ先": str(other / CL_DIFF)})
        code, out = run("restore", "-d", self.dfile, str(self.rec(CL_DIFF).no))
        self.assertEqual(code, 1)
        self.assertIn("納品コピーのバックアップではない", out)
        self.assertEqual((self.dest / CL_DIFF).read_bytes(), b"dummy")

    # ---- 設定
    def test_without_dest_root(self):
        text = (self.demo / "delivery.toml").read_text(encoding="utf-8").replace('dest_root = "dest"\n', "")
        (self.demo / "delivery.toml").write_text(text, encoding="utf-8")
        code, out = self.deliver("list")
        self.assertEqual(code, 2)
        self.assertIn("dest_root", out)
        code, out = run("report", "-d", self.dfile)  # report はこれまでどおり動き、納品先のチェックをしない
        self.assertEqual(code, 0, out)
        self.assertNotIn("管理表にないファイル", self.cats())

    def test_dest_folder_not_created_yet(self):
        for p in sorted(self.dest.rglob("*"), reverse=True):
            p.unlink() if p.is_file() else p.rmdir()
        self.dest.rmdir()
        run("report", "-d", self.dfile)
        self.assertIn("納品先のフォルダなし", self.cats())
        self.deliver("check")
        self.assertTrue((self.dest / CL_NEW).exists())

    # ---- 納品先のチェック
    def test_stray_file_and_ignored_temp(self):
        msgs = [i[3] for i in self.issues() if i[2] == "管理表にないファイル"]
        self.assertEqual(len(msgs), 1)
        self.assertIn("古い資料.xlsx", msgs[0])
        self.assertFalse(any("~$" in i[3] for i in self.issues()))

    def test_leftover_temp_file_warning(self):
        touch(self.dest / (CL_NEW + fileops.TMP_SUFFIX))
        run("report", "-d", self.dfile)
        self.assertIn("一時ファイルの残り", self.cats())

    def test_names_compared_by_nfc_and_case(self):
        self.assertEqual(norm_key(unicodedata.normalize("NFD", "案件A/090.SD消化/ガイド_A.xlsx")),
                         norm_key("案件A\\090.SD消化\\ガイド_a.XLSX"))
        # 同じ行の納品先に、大文字・小文字と文字の表し方だけが違う名前で置いても「管理表にない」にならない
        rel = unicodedata.normalize("NFD", CL_NEW.replace("_B.xlsx", "_b.XLSX"))
        touch(self.dest / rel)
        run("report", "-d", self.dfile)
        strays = [i[3] for i in self.issues() if i[2] == "管理表にないファイル"]
        self.assertEqual(len(strays), 1, strays)  # demo の「古い資料.xlsx」だけ

    def test_source_updated_after_delivery(self):
        (self.dest / CL_DIFF).unlink()
        self.deliver("check")
        Path(self.rec(CL_NEW).src).write_bytes(b"stamped again")
        run("report", "-d", self.dfile)
        self.assertIn("納品後に原本が更新", self.cats())

    def test_destination_changed_and_hash_setting(self):
        (self.dest / CL_DIFF).unlink()
        self.deliver("check")
        (self.dest / CL_NEW).write_bytes(b"DUMMY")  # 同じサイズで中身だけ違う
        run("report", "-d", self.dfile)
        self.assertIn("納品先の内容違い", self.cats())
        # 照合をやめると、サイズしか見ないので気づかない
        cfg = self.demo / "config.toml"
        cfg.write_text(cfg.read_text(encoding="utf-8").replace("verify_hash_on_report = true",
                                                               "verify_hash_on_report = false"), encoding="utf-8")
        run("report", "-d", self.dfile)
        self.assertNotIn("納品先の内容違い", self.cats())
        (self.dest / CL_NEW).write_bytes(b"much longer content")
        run("report", "-d", self.dfile)
        self.assertIn("納品先の内容違い", self.cats())
        (self.dest / CL_NEW).unlink()
        run("report", "-d", self.dfile)
        self.assertIn("納品先なし", self.cats())

    def test_bad_dest_rel_is_skipped(self):
        self.edit(cl_name("030", "社内", 1, "B"), **{"納品先(相対)": "案件B/../../outside.xlsx"})
        code, out = self.deliver("list")
        self.assertIn("納品先(相対)が不正", out)
        self.assertNotIn("納品先が不正", self.cats())  # list は管理表を書き換えない
        run("report", "-d", self.dfile)
        self.assertIn("納品先が不正", self.cats())
        (self.dest / CL_DIFF).unlink()
        self.deliver("check")
        self.assertFalse((self.demo / "outside.xlsx").exists())

    # ---- 手動・納品先が空
    def test_auto_row_without_dest_shows_candidate(self):
        self.edit(MISNAMED, 対象="自動")
        code, out = self.deliver("list")
        self.assertIn("納品先(相対)が空（自動）", out)
        self.assertIn("候補: 案件A/040.設計/成果物/レビュー記録表_設計_A.xlsx", out)

    def test_manual_row(self):
        self.edit(MISNAMED, 対象="手動")
        code, out = self.deliver("list")
        self.assertIn("手で配置（手動）: 1 件", out)
        self.assertIn("候補: 案件A/040.設計/成果物/レビュー記録表_設計_A.xlsx", out)
        # 納品先(相対)を入れずに手で置くと、案内付きで「管理表にないファイル」になる
        placed = self.dest / "案件A" / "040.設計" / "成果物" / MISNAMED
        touch(placed, b"wrong version")
        run("report", "-d", self.dfile)
        hit = [i for i in self.issues() if i[2] == "管理表にないファイル" and MISNAMED in i[3]]
        self.assertEqual(len(hit), 1)
        self.assertIn("同じ名前の行の納品先(相対)が空です", hit[0][3])
        self.assertEqual(hit[0][4], str(self.rec(filename=MISNAMED).no))
        # 納品先(相対)を入れて「済」にすると、中身を原本と比べる
        self.edit(MISNAMED, **{"納品先(相対)": "案件A/040.設計/成果物/" + MISNAMED, "納品コピー": "済"})
        code, out = self.deliver("list")
        self.assertIn(f"{MISNAMED}  納品先: 内容が違う", out)
        run("report", "-d", self.dfile)
        self.assertIn("手動配置の内容違い", self.cats())
        self.assertFalse(any(MISNAMED in i[3] for i in self.issues() if i[2] == "管理表にないファイル"))
        placed.write_bytes(b"dummy")
        run("report", "-d", self.dfile)
        self.assertNotIn("手動配置の内容違い", self.cats())
        # 手動の行はツールがコピー・上書きしない
        placed.write_bytes(b"wrong version")
        self.deliver("overwrite")
        self.assertEqual(placed.read_bytes(), b"wrong version")


if __name__ == "__main__":
    unittest.main()
