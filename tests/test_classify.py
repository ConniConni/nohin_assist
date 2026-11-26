import unittest
from pathlib import Path

from nohin.classify import analyze_location, build_record, classify_name
from nohin.config import load_config
from nohin.model import KIND_CL, KIND_DOC, KIND_OTHER, KIND_RV, ST_NA, ST_TODO, TARGET_AUTO, TARGET_CHECK, TARGET_EXCLUDE

ROOT = Path(__file__).resolve().parent.parent
CFG = load_config(ROOT / "config.example.toml")


class ClassifyNameTest(unittest.TestCase):
    def test_documents(self):
        for name, doc, code in [("機能設計書_xxx.xlsx", "機能設計書", "040"),
                                ("調査検討書_a_b.xlsx", "調査検討書", "030"),
                                ("調査_xxx.xlsx", "調査検討書", "030"),
                                ("試験結果報告書_x.xlsx", "試験結果報告書", "090")]:
            info = classify_name(name, CFG)
            self.assertEqual(info.kind, KIND_DOC, name)
            self.assertEqual(info.doc_rule.name, doc)
            self.assertEqual(info.name_code, code)

    def test_review_record_half_and_full_width(self):
        for name in ["レビュー記録表(社外)(2回目)_設計_xxx.xlsx", "レビュー記録表（社外）（2回目）_設計_xxx.xlsx"]:
            info = classify_name(name, CFG)
            self.assertEqual((info.kind, info.side, info.round, info.name_code), (KIND_RV, "社外", 2, "040"), name)

    def test_review_record_process_names(self):
        self.assertEqual(classify_name("レビュー記録表(社内)(1回目)_UD作成_x.xlsx", CFG).name_code, "060")
        self.assertEqual(classify_name("レビュー記録表(社内)(1回目)_試験結果報告書_x.xlsx", CFG).name_code, "090")
        info = classify_name("レビュー記録表(社内)(1回目)_存在しない_x.xlsx", CFG)
        self.assertEqual(info.kind, KIND_RV)
        self.assertIsNone(info.name_code)
        self.assertTrue(info.problems)

    def test_checklist(self):
        info = classify_name("060_レビューチェックリスト_UD仕様書(社内)(3回目)_..._.xlsx", CFG)
        self.assertEqual((info.kind, info.name_code, info.side, info.round), (KIND_CL, "060", "社内", 3))
        self.assertFalse(info.problems)
        bad = classify_name("060_レビューチェックリスト_UD作成(社内)(1回目)_x.xlsx", CFG)  # 記録表側の工程名
        self.assertTrue(bad.problems)

    def test_other(self):
        for name in ["メモ.txt", "設計書_x.xlsx", "レビュー記録表_x.xlsx", "機能設計書.xlsx"]:
            self.assertEqual(classify_name(name, CFG).kind, KIND_OTHER, name)


class LocationTest(unittest.TestCase):
    def test_patterns(self):
        self.assertEqual(analyze_location(("a.xlsx",), CFG).kind, "doc")
        loc = analyze_location(("成果物", "内部Rv", "a.xlsx"), CFG)
        self.assertEqual((loc.kind, loc.side_dir), ("review", "社内"))
        loc = analyze_location(("成果物", "外部Rv", "20261120", "a.xlsx"), CFG)
        self.assertEqual((loc.kind, loc.side_dir, loc.date_dir), ("review", "社外", "20261120"))
        self.assertEqual(analyze_location(("成果物", "a.xlsx"), CFG).kind, "other")
        self.assertEqual(analyze_location(("成果物", "外部Rv", "old", "a.xlsx"), CFG).kind, "other")
        self.assertEqual(analyze_location(("成果物", "外部Rv", "20261120", "x", "a.xlsx"), CFG).kind, "other")


class BuildRecordTest(unittest.TestCase):
    def rec(self, rel):
        return build_record("案件A", rel, "/x/" + rel, CFG, "2026-01-01 00:00")

    def test_doc_with_partner(self):
        r = self.rec("040.設計/機能設計書_x.xlsx")
        self.assertEqual((r.process, r.kind, r.target), ("040", KIND_DOC, TARGET_AUTO))
        self.assertEqual((r.st_self, r.st_send, r.st_partner, r.st_copy), (ST_TODO, ST_TODO, ST_TODO, ST_TODO))
        self.assertEqual(r.dest_rel, "案件A/040.設計/機能設計書_x.xlsx")
        self.assertEqual(r.teams_rel, r.dest_rel)

    def test_doc_self_only(self):
        r = self.rec("060.UD作成/単体試験仕様書_x.xlsx")
        self.assertEqual((r.st_self, r.st_send), (ST_TODO, ST_NA))
        self.assertEqual(r.teams_rel, "")

    def test_research_doc_no_stamp(self):
        r = self.rec("030.調査/調査検討書_x.xlsx")
        self.assertEqual((r.st_self, r.st_send, r.st_copy), (ST_NA, ST_NA, ST_TODO))

    def test_review_flattened_dest(self):
        r = self.rec("040.設計/成果物/外部Rv/20261120/レビュー記録表(社外)(1回目)_設計_x.xlsx")
        self.assertEqual(r.dest_rel, "案件A/040.設計/成果物/レビュー記録表(社外)(1回目)_設計_x.xlsx")
        self.assertEqual((r.st_self, r.st_send), (ST_TODO, ST_TODO))
        r = self.rec("040.設計/成果物/内部Rv/レビュー記録表(社内)(1回目)_設計_x.xlsx")
        self.assertEqual((r.st_self, r.st_send), (ST_TODO, ST_NA))

    def test_checklist_no_stamp(self):
        r = self.rec("040.設計/成果物/外部Rv/040_レビューチェックリスト_設計(社外)(1回目)_x.xlsx")
        self.assertEqual((r.st_self, r.st_send, r.st_copy), (ST_NA, ST_NA, ST_TODO))

    def test_other_files_excluded(self):
        for rel in ["040.設計/メモ.txt", "040.設計/成果物/内部Rv/議事録.docx", "040.設計/作業用/下書き.xlsx"]:
            self.assertEqual(self.rec(rel).target, TARGET_EXCLUDE, rel)

    def test_dest_uses_configured_folder_name(self):
        r = self.rec("040.設計_旧/機能設計書_x.xlsx")
        self.assertEqual(r.dest_rel, "案件A/040.設計/機能設計書_x.xlsx")

    def test_needs_check(self):
        for rel in ["040.設計/成果物/内部Rv/レビュー記録表_x.xlsx",                         # 名前の誤り（置き場所は納品物）
                    "040.設計/成果物/機能設計書_x.xlsx",                                    # 置き場所違い
                    "040.設計/成果物/外部Rv/レビュー記録表(社内)(1回目)_設計_x.xlsx",       # 社内外違い
                    "060.UD作成/機能設計書_x.xlsx",                                        # 工程違い
                    "040.設計/成果物/内部Rv/030_レビューチェックリスト_調査(社内)(1回目)_x.xlsx"]:
            self.assertEqual(self.rec(rel).target, TARGET_CHECK, rel)


if __name__ == "__main__":
    unittest.main()
