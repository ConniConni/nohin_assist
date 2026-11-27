"""ダミーのプロジェクトルートを作る（テストと動作確認用）。

    python -m tests.dummy demo        # demo/ に3案件分と delivery.toml を作る
"""

from __future__ import annotations

import shutil
import sys
from datetime import date
from pathlib import Path

from openpyxl import Workbook

ROOT = Path(__file__).resolve().parent.parent

FOLDERS = {"030": "030.調査", "040": "040.設計", "050": "050.製造", "060": "060.UD作成",
           "070": "070.UD消化", "080": "080.SD作成", "090": "090.SD消化"}
CL = {"030": "調査", "040": "設計", "050": "CD", "060": "UD仕様書", "070": "UD結果", "080": "SD仕様書",
      "090": "試験結果報告書"}
RV = {"030": "調査", "040": "設計", "050": "CD", "060": "UD作成", "070": "UD消化", "080": "SD作成",
      "090": "試験結果報告書"}
DOCS = {"030": ["調査検討書"], "040": ["機能設計書"], "050": [], "060": ["単体試験仕様書"],
        "070": ["単体試験成績書"], "080": ["結合試験仕様書"], "090": ["結合試験成績書", "試験結果報告書"]}
# 工程ごとのレビュー（社内外, 回数）
REVIEWS = {"030": [("社内", 1), ("社外", 1)], "040": [("社内", 1), ("社外", 1)], "050": [("社内", 1)],
           "060": [("社内", 1)], "070": [("社内", 1)], "080": [("社内", 1), ("社外", 1)],
           "090": [("社内", 1), ("社外", 1)]}


def touch(p: Path) -> Path:
    p.parent.mkdir(parents=True, exist_ok=True)
    p.write_bytes(b"dummy")
    return p


# レビュー記録表の指摘: (指摘事項, 処置欄, 処置の日付, 捺印(自社)欄の文字)
COMPLETE = [("誤字があります", "修正しました", date(2026, 11, 18), ""),
            ("条件分岐の記載漏れ", "追記しました", date(2026, 11, 20), "")]
HEADERS = ["No.", "日付", "指摘事項", "指摘者", "対処欄", "処置欄", "日付", "捺印(自社)", "担当者", "捺印(取引先)"]


def write_review_record(p: Path, findings=COMPLETE, *, extra_upper_rows: int = 0, with_table: bool = True,
                        with_approval: bool = True, header_merged: bool = False) -> Path:
    """文字だけが入った本物のレビュー記録表（xlsx）を作る。印影はまだ入れない。

    上部: タイトル、承認欄（H2「承認」の下の H3:H4）、実施日時・場所・参加者・ドキュメント名。
    指摘の表: 見出し（A〜J列）、指摘の行、No. と「-」だけの行が2行、空行のあとに備考（表の外）。
    """
    p.parent.mkdir(parents=True, exist_ok=True)
    wb = Workbook()
    ws = wb.active
    ws.title = "レビュー記録表"
    ws["A1"] = "レビュー記録表"
    if with_approval:
        ws["H2"], ws["I2"], ws["J2"] = "承認", "確認", "作成"
        for col in "HIJ":
            ws.merge_cells(f"{col}3:{col}4")
    r = 3
    for label in ["レビュー実施日時", "場所", "参加者", "ドキュメント名"]:
        ws.cell(r, 1, label)
        r += 1
    for i in range(1 + extra_upper_rows):  # ドキュメントが多いと上部の行が増える
        ws.cell(r, 2, f"ドキュメント{i + 1}")
        r += 1
    h = r + 1
    if not with_table:
        ws.cell(h, 1, "（指摘の表なし）")
        wb.save(p)
        return p
    for c, text in enumerate(HEADERS, start=1):
        ws.cell(h, c, text)
        if header_merged:
            ws.merge_cells(start_row=h, start_column=c, end_row=h + 1, end_column=c)
    r = h + (2 if header_merged else 1)
    for n, (finding, action, d, mark) in enumerate(findings, start=1):
        for c, v in enumerate([n, date(2026, 11, 10), finding, "山田", "対応方針", action, d, mark, "", ""], start=1):
            if v not in ("", None):
                ws.cell(r, c, v)
        r += 1
    for n in range(len(findings) + 1, len(findings) + 3):  # 様式に最初からある空き行
        ws.cell(r, 1, n)
        ws.cell(r, 8, "-")
        r += 1
    ws.cell(r + 6, 3, "備考: 特になし")  # 空行が続いたあとの表の外（指摘として数えない）
    wb.save(p)
    return p


def rv_name(code, side, n, tag="xxx"):
    return f"レビュー記録表({side})({n}回目)_{RV[code]}_{tag}.xlsx"


def cl_name(code, side, n, tag="xxx"):
    return f"{code}_レビューチェックリスト_{CL[code]}({side})({n}回目)_{tag}.xlsx"


def make_clean_project(root: Path, codes=tuple(FOLDERS), date_dirs: bool = False, tag: str = "xxx") -> Path:
    """規則どおりのプロジェクトを作る。date_dirs=True ならパターン2（日付フォルダ）。"""
    for code in codes:
        base = root / FOLDERS[code]
        for d in DOCS[code]:
            touch(base / f"{d}_{tag}.xlsx")
        (base / "back").mkdir(parents=True, exist_ok=True)
        for side, n in REVIEWS[code]:
            sub = base / "成果物" / ("内部Rv" if side == "社内" else "外部Rv")
            if date_dirs:
                sub = sub / f"2026110{n}"
            write_review_record(sub / rv_name(code, side, n, tag))
            touch(sub / cl_name(code, side, n, tag))
    return root


def make_demo(out: Path) -> Path:
    if out.exists():
        shutil.rmtree(out)
    a = make_clean_project(out / "src" / "案件A", tag="A")
    # 案件A: 例外をいくつか入れる
    touch(a / "040.設計" / "back" / "機能設計書_A_旧版.xlsx")                          # back は対象外
    write_review_record(a / "040.設計" / "成果物" / "内部Rv" / rv_name("040", "社内", 2, "A"))  # 2回目（チェックリストなし）
    touch(a / "040.設計" / "メモ.txt")                                                 # 規則外（対象外）
    touch(a / "040.設計" / "成果物" / "内部Rv" / "レビュー記録表_設計_A.xlsx")          # 規則外（名前の誤り→要確認）
    write_review_record(a / "060.UD作成" / "成果物" / "外部Rv" / rv_name("060", "社内", 1, "A2"))  # 社内外の不一致
    touch(a / "080.SD作成" / "成果物" / "内部Rv" / "080_レビューチェックリスト_SD仕様書（社内）（2回目）_A.xlsx")  # 全角括弧
    touch(a / "040.設計" / "~$機能設計書_A.xlsx")                                       # Excel の一時ファイル（無視）
    touch(a / "040.設計" / "作業用" / "機能設計書_A_作業中.xlsx")                        # 直下以外のフォルダ（載せない）
    touch(a / "040.設計" / "成果物" / "外部Rv" / "機能設計書_A.xlsx")                    # レビューフォルダのドキュメントの写し（載せない）
    write_review_record(a / "040.設計" / "成果物" / "内部Rv" / "過去分" / rv_name("040", "社内", 1, "A_old"))  # 日付でないフォルダ（載せない）
    touch(a / "040.設計" / "機能設計書_A_画面2.xlsx")                                   # 複数あってよい（allow_multiple）

    # 案件A: レビュー記録表の中身
    rv = lambda code, side, n: a / FOLDERS[code] / "成果物" / ("内部Rv" if side == "社内" else "外部Rv") / rv_name(code, side, n, "A")  # noqa: E731
    write_review_record(rv("040", "社内", 1), extra_upper_rows=3, header_merged=True)       # 上部の行が多い・見出しが2行結合
    write_review_record(rv("040", "社外", 1), [COMPLETE[0], ("説明不足", "", None, "")])     # 処置中
    write_review_record(rv("030", "社内", 1), [("誤記", "修正済み", None, "")])             # 日付なし
    write_review_record(rv("030", "社外", 1), [("誤記", "修正済み", "11/20", "")])          # 日付が読めない（年がない）
    write_review_record(rv("050", "社内", 1), [COMPLETE[0], ("質問", "回答のみ", None, "－"),
                                               ("確認事項", "", None, "-")])                # 捺印不要（全角・半角）
    write_review_record(rv("070", "社内", 1), [("指摘なし", "", None, "")])                 # 指摘なしの印
    write_review_record(rv("080", "社内", 1), [COMPLETE[0], ("", "追記しました", date(2026, 11, 21), "")])  # 指摘欄が空の行に記入
    write_review_record(rv("090", "社内", 1), with_table=False)                             # 表なし
    write_review_record(rv("090", "社外", 1), [])                                           # 指摘0件

    b = make_clean_project(out / "src" / "案件B", codes=("030", "040"), date_dirs=True, tag="B")
    write_review_record(b / "040.設計" / "成果物" / "外部Rv" / "20261120" / rv_name("040", "社外", 3, "B"))  # 2回目が抜け
    touch(b / "040.設計" / "成果物" / "外部Rv" / "20261120" / cl_name("040", "社外", 3, "B"))
    touch(b / "060.UD作成" / "単体試験仕様書_B.xlsx")                                     # 対象工程外（読み飛ばし）

    c = make_clean_project(out / "src" / "案件C", codes=("060", "070", "080", "090"), tag="C")
    (c / "090.SD消化" / "試験結果報告書_C.xlsx").unlink()                                  # ドキュメントの欠け
    touch(c / "080.SD作成" / "結合試験仕様書_C_v2.xlsx")                                   # ドキュメント重複
    write_review_record(c / "060.UD作成" / "成果物" / "内部Rv" / rv_name("060", "社内", 1, "C"),
                        with_approval=False)                                                  # 承認欄なし

    shutil.copy(ROOT / "config.example.toml", out / "config.toml")
    (out / "delivery.toml").write_text(
        '[delivery]\nname = "デモ納品"\nwork_dir = "work"\n\n'
        '[[projects]]\nname = "案件A"\nroot = "src/案件A"\nallow_multiple = ["機能設計書"]\n\n'
        '[[projects]]\nname = "案件B"\nroot = "src/案件B"\nprocesses = ["030", "040"]\n\n'
        '[[projects]]\nname = "案件C"\nroot = "src/案件C"\nprocesses = ["060", "070", "080", "090"]\n',
        encoding="utf-8")
    return out


if __name__ == "__main__":
    target = Path(sys.argv[1] if len(sys.argv) > 1 else "demo").resolve()
    make_demo(target)
    print(f"ダミーを作成しました: {target}")
    print(f"  python -m nohin collect -d {target / 'delivery.toml'}")
