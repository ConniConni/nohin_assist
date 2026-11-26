"""ダミーのプロジェクトルートを作る（テストと動作確認用）。

    python -m tests.dummy demo        # demo/ に3案件分と delivery.toml を作る
"""

from __future__ import annotations

import shutil
import sys
from pathlib import Path

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
            touch(sub / rv_name(code, side, n, tag))
            touch(sub / cl_name(code, side, n, tag))
    return root


def make_demo(out: Path) -> Path:
    if out.exists():
        shutil.rmtree(out)
    a = make_clean_project(out / "src" / "案件A", tag="A")
    # 案件A: 例外をいくつか入れる
    touch(a / "040.設計" / "back" / "機能設計書_A_旧版.xlsx")                          # back は対象外
    touch(a / "040.設計" / "成果物" / "内部Rv" / rv_name("040", "社内", 2, "A"))       # 2回目（チェックリストなし）
    touch(a / "040.設計" / "メモ.txt")                                                 # 規則外（対象外）
    touch(a / "040.設計" / "成果物" / "内部Rv" / "レビュー記録表_設計_A.xlsx")          # 規則外（名前の誤り→要確認）
    touch(a / "060.UD作成" / "成果物" / "外部Rv" / rv_name("060", "社内", 1, "A2"))   # 社内外の不一致
    touch(a / "080.SD作成" / "成果物" / "内部Rv" / "080_レビューチェックリスト_SD仕様書（社内）（2回目）_A.xlsx")  # 全角括弧
    touch(a / "040.設計" / "~$機能設計書_A.xlsx")                                       # Excel の一時ファイル（無視）

    b = make_clean_project(out / "src" / "案件B", codes=("030", "040"), date_dirs=True, tag="B")
    touch(b / "040.設計" / "成果物" / "外部Rv" / "20261120" / rv_name("040", "社外", 3, "B"))  # 2回目が抜け
    touch(b / "040.設計" / "成果物" / "外部Rv" / "20261120" / cl_name("040", "社外", 3, "B"))
    touch(b / "060.UD作成" / "単体試験仕様書_B.xlsx")                                     # 対象工程外（読み飛ばし）

    c = make_clean_project(out / "src" / "案件C", codes=("060", "070", "080", "090"), tag="C")
    (c / "090.SD消化" / "試験結果報告書_C.xlsx").unlink()                                  # ドキュメントの欠け
    touch(c / "080.SD作成" / "結合試験仕様書_C_v2.xlsx")                                   # ドキュメント重複

    shutil.copy(ROOT / "config.example.toml", out / "config.toml")
    (out / "delivery.toml").write_text(
        '[delivery]\nname = "デモ納品"\nwork_dir = "work"\n\n'
        '[[projects]]\nname = "案件A"\nroot = "src/案件A"\n\n'
        '[[projects]]\nname = "案件B"\nroot = "src/案件B"\nprocesses = ["030", "040"]\n\n'
        '[[projects]]\nname = "案件C"\nroot = "src/案件C"\nprocesses = ["060", "070", "080", "090"]\n',
        encoding="utf-8")
    return out


if __name__ == "__main__":
    target = Path(sys.argv[1] if len(sys.argv) > 1 else "demo").resolve()
    make_demo(target)
    print(f"ダミーを作成しました: {target}")
    print(f"  python -m nohin collect -d {target / 'delivery.toml'}")
