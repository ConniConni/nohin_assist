"""デモデータの正誤表と、実際の出力（管理表）との突き合わせ。

    python -m tests.verify_demo                         # demo/ の管理表を照合して ○/× を表示（collect のあと）
    python -m tests.verify_demo --stage deliver         # deliver（overwrite）のあとの状態を照合
    python -m tests.verify_demo demo --markdown docs/デモ正誤表.md   # 正誤表を Markdown で書き出す

正誤表（CASES / EXPECTED_ISSUES / EXPECTED_TOTALS / DELIVER_CASES / EXPECTED_COPIED）は
tests/dummy.py の make_demo() と対応させる。
ダミーの内容を変えたら、ここも合わせて直すこと。
"""

from __future__ import annotations

import argparse
import sys
from dataclasses import dataclass, field
from pathlib import Path

from openpyxl import load_workbook

from nohin.config import load_config, load_delivery
from nohin.manifest import read_manifest

# ステータスの並び: (指摘展開, 自社捺印, 取引先送付, 取引先捺印, 納品コピー)
NEED_ALL = ("不要", "未", "未", "未", "未")   # ドキュメント（自社＋取引先）
RV_EXT = ("未", "未", "未", "未", "未")       # 社外レビュー記録表（指摘展開あり）
SELF_ONLY = ("不要", "未", "不要", "不要", "未")
COPY_ONLY = ("不要", "不要", "不要", "不要", "未")


@dataclass
class Case:
    id: str
    group: str          # 仕込み / 正常
    project: str
    rel: str            # プロジェクトルートからの相対パス
    setup: str          # 何を仕込んだか
    present: bool = True            # 管理表に載るか
    target: str | None = None       # 期待する「対象」
    statuses: tuple[str, ...] | None = None
    dest_rel: str | None = None     # 期待する納品先（相対）
    rv: tuple[str, str, str] | None = None  # 期待する (記録表の状況, 指摘の進捗, 承認印の日付)
    issues: list[tuple[str, str]] = field(default_factory=list)  # 期待する (区分, 種類)

    def expectation(self) -> str:
        if not self.present:
            return "管理表に載らない"
        parts = [f"対象={self.target}"] if self.target else []
        if self.statuses:
            parts.append("ステータス=" + "/".join(self.statuses))
        if self.dest_rel:
            parts.append(f"納品先={self.dest_rel}")
        if self.rv:
            parts.append(f"状況={self.rv[0]}、進捗={self.rv[1]}、承認印の日付={self.rv[2] or '（空）'}")
        parts += [f"{lv}「{cat}」" for lv, cat in self.issues]
        return "、".join(parts)


CASES: list[Case] = [
    # ---- あえて不正・例外を仕込んだもの ----
    Case("A-1", "仕込み", "案件A", "040.設計/back/機能設計書_A_旧版.xlsx",
         "back/ の下に旧版を置いた", present=False),
    Case("A-2", "仕込み", "案件A", "040.設計/~$機能設計書_A.xlsx",
         "Excel の一時ファイルを置いた", present=False),
    Case("A-3", "仕込み", "案件A", "040.設計/成果物/内部Rv/レビュー記録表(社内)(2回目)_設計_A.xlsx",
         "2回目のレビュー記録表だけ置き、チェックリストを置かない",
         target="自動", statuses=SELF_ONLY, issues=[("警告", "ペア不足")]),
    Case("A-4", "仕込み", "案件A", "040.設計/メモ.txt",
         "命名規則外のテキストファイルを工程直下に置いた",
         target="除外", issues=[("情報", "規則外（対象外）")]),
    Case("A-5", "仕込み", "案件A", "040.設計/成果物/内部Rv/レビュー記録表_設計_A.xlsx",
         "レビューフォルダに名前を間違えた Excel を置いた",
         target="要確認", issues=[("警告", "規則外")]),
    Case("A-6", "仕込み", "案件A", "060.UD作成/成果物/外部Rv/レビュー記録表(社内)(1回目)_UD作成_A2.xlsx",
         "外部Rv フォルダに「社内」のレビュー記録表を置いた",
         target="要確認", issues=[("エラー", "社内外不一致")]),
    Case("A-7", "仕込み", "案件A", "080.SD作成/成果物/内部Rv/080_レビューチェックリスト_SD仕様書（社内）（2回目）_A.xlsx",
         "全角括弧のチェックリスト（2回目、対応するレビュー記録表なし）",
         target="自動", statuses=COPY_ONLY, issues=[("情報", "全角括弧"), ("警告", "ペア不足")]),
    Case("B-1", "仕込み", "案件B", "040.設計/成果物/外部Rv/20261120/レビュー記録表(社外)(3回目)_設計_B.xlsx",
         "社外レビューを1回目・3回目だけにした（2回目が抜け）",
         target="自動", statuses=RV_EXT, issues=[("警告", "回数の抜け")]),
    Case("B-2", "仕込み", "案件B", "060.UD作成/単体試験仕様書_B.xlsx",
         "対象工程（030・040）以外の工程フォルダにファイルを置いた", present=False),
    Case("C-1", "仕込み", "案件C", "080.SD作成/結合試験仕様書_C.xlsx",
         "同じ工程に結合試験仕様書を2つ置いた（_C と _C_v2）",
         target="自動", issues=[("エラー", "ドキュメント重複")]),
    Case("C-2", "仕込み", "案件C", "080.SD作成/結合試験仕様書_C_v2.xlsx",
         "（C-1 の2つ目）", target="自動", issues=[("エラー", "ドキュメント重複")]),
    Case("C-3", "仕込み", "案件C", "090.SD消化/試験結果報告書_C.xlsx",
         "試験結果報告書を削除した", present=False, issues=[("警告", "ドキュメントなし")]),
    # ---- 正常系（種別ごとのステータス・納品先） ----
    Case("N-1", "正常", "案件A", "030.調査/調査検討書_A.xlsx", "調査検討書（捺印なし）",
         target="自動", statuses=COPY_ONLY, dest_rel="案件A/030.調査/調査検討書_A.xlsx"),
    Case("N-2", "正常", "案件A", "040.設計/機能設計書_A.xlsx", "機能設計書（自社＋取引先）",
         target="自動", statuses=NEED_ALL, dest_rel="案件A/040.設計/機能設計書_A.xlsx"),
    Case("N-3", "正常", "案件A", "060.UD作成/単体試験仕様書_A.xlsx", "単体試験仕様書（自社のみ）",
         target="自動", statuses=SELF_ONLY),
    Case("N-4", "正常", "案件A", "040.設計/成果物/外部Rv/レビュー記録表(社外)(1回目)_設計_A.xlsx",
         "社外レビュー記録表（自社＋取引先）・処置中（1件だけ処置済み）", target="自動", statuses=RV_EXT,
         dest_rel="案件A/040.設計/成果物/レビュー記録表(社外)(1回目)_設計_A.xlsx",
         rv=("処置中", "指摘2・処置1・日付1", "")),
    Case("N-5", "正常", "案件A", "040.設計/成果物/内部Rv/レビュー記録表(社内)(1回目)_設計_A.xlsx",
         "社内レビュー記録表（自社のみ）・上部の行が多い・見出しが2行結合", target="自動", statuses=SELF_ONLY,
         rv=("処置完了", "指摘2・処置2・日付2", "2026-11-20")),
    Case("N-6", "正常", "案件A", "040.設計/成果物/外部Rv/040_レビューチェックリスト_設計(社外)(1回目)_A.xlsx",
         "チェックリスト（捺印なし）", target="自動", statuses=COPY_ONLY),
    Case("N-7", "正常", "案件B", "040.設計/成果物/外部Rv/20261101/レビュー記録表(社外)(1回目)_設計_B.xlsx",
         "日付フォルダの中（パターン2）・完成した社外の記録表", target="自動", statuses=RV_EXT,
         dest_rel="案件B/040.設計/成果物/レビュー記録表(社外)(1回目)_設計_B.xlsx",
         rv=("処置完了", "指摘2・処置2・日付2", "")),
    Case("N-8", "正常", "案件C", "090.SD消化/結合試験成績書_C.xlsx", "結合試験成績書（自社＋取引先）",
         target="自動", statuses=NEED_ALL),
    Case("N-9", "正常", "案件C", "070.UD消化/成果物/内部Rv/レビュー記録表(社内)(1回目)_UD消化_C.xlsx",
         "完成した社内の記録表（表の下の空行のあとに備考あり）", target="自動", statuses=SELF_ONLY,
         rv=("処置完了", "指摘2・処置2・日付2", "2026-11-20")),
    # ---- レビュー記録表の中身に仕込んだもの ----
    Case("R-1", "仕込み", "案件A", "030.調査/成果物/内部Rv/レビュー記録表(社内)(1回目)_調査_A.xlsx",
         "処置欄はあるが処置の日付が空", target="自動",
         rv=("処置中", "指摘1・処置1・日付0", ""), issues=[("警告", "日付なし")]),
    Case("R-2", "仕込み", "案件A", "030.調査/成果物/外部Rv/レビュー記録表(社外)(1回目)_調査_A.xlsx",
         "処置の日付が「11/20」（年がなく読めない）", target="自動",
         rv=("処置中", "指摘1・処置1・日付0", ""), issues=[("警告", "日付なし")]),
    Case("R-3", "仕込み", "案件A", "050.製造/成果物/内部Rv/レビュー記録表(社内)(1回目)_CD_A.xlsx",
         "捺印欄が全角「－」・半角「-」の行（捺印不要。処置・日付なしでも可）", target="自動",
         rv=("処置完了", "指摘3（捺印不要2）・処置2・日付1", "2026-11-18")),
    Case("R-4", "仕込み", "案件A", "070.UD消化/成果物/内部Rv/レビュー記録表(社内)(1回目)_UD消化_A.xlsx",
         "指摘事項に「指摘なし」とだけ書いた", target="自動",
         rv=("指摘なし", "指摘0・処置0・日付0", ""), issues=[("情報", "指摘0件")]),
    Case("R-5", "仕込み", "案件A", "080.SD作成/成果物/内部Rv/レビュー記録表(社内)(1回目)_SD作成_A.xlsx",
         "指摘事項が空の行に処置と日付を書いた", target="自動",
         rv=("処置完了", "指摘1・処置1・日付1", "2026-11-18"), issues=[("警告", "指摘欄が空の行に記入")]),
    Case("R-6", "仕込み", "案件A", "090.SD消化/成果物/内部Rv/レビュー記録表(社内)(1回目)_試験結果報告書_A.xlsx",
         "指摘の表がない", target="自動", rv=("表なし", "", ""), issues=[("警告", "指摘表なし")]),
    Case("R-7", "仕込み", "案件A", "090.SD消化/成果物/外部Rv/レビュー記録表(社外)(1回目)_試験結果報告書_A.xlsx",
         "指摘の表はあるが指摘の行がない", target="自動",
         rv=("指摘なし", "指摘0・処置0・日付0", ""), issues=[("情報", "指摘0件")]),
    Case("R-8", "仕込み", "案件C", "060.UD作成/成果物/内部Rv/レビュー記録表(社内)(1回目)_UD作成_C.xlsx",
         "上部に承認欄がない", target="自動",
         rv=("処置完了", "指摘2・処置2・日付2", "2026-11-20"), issues=[("警告", "承認欄なし")]),
]

# 警告シートに出るべきもの（これ以外が出たら×）: (区分, 案件, 種類, 内容に含まれる文字列)
EXPECTED_ISSUES: list[tuple[str, str, str, str]] = [
    ("エラー", "案件A", "社内外不一致", "レビュー記録表(社内)(1回目)_UD作成_A2.xlsx"),
    ("エラー", "案件C", "ドキュメント重複", "結合試験仕様書"),
    ("警告", "案件A", "ペア不足", "工程 040 社内 2回目"),
    ("警告", "案件A", "ペア不足", "工程 080 社内 2回目"),
    ("警告", "案件A", "規則外", "レビュー記録表_設計_A.xlsx"),
    ("警告", "案件B", "回数の抜け", "工程 040 社外"),
    ("警告", "案件C", "ドキュメントなし", "試験結果報告書"),
    ("警告", "案件A", "日付なし", "No.1（空）"),
    ("警告", "案件A", "日付なし", "「11/20」を日付として読めません"),
    ("警告", "案件A", "指摘欄が空の行に記入", "SD作成_A.xlsx"),
    ("警告", "案件A", "指摘表なし", "試験結果報告書_A.xlsx"),
    ("警告", "案件C", "承認欄なし", "UD作成_C.xlsx"),
    ("警告", "案件A", "管理表にないファイル", "案件A/040.設計/古い資料.xlsx"),
    ("情報", "案件A", "全角括弧", "SD仕様書（社内）（2回目）"),
    ("情報", "案件A", "規則外（対象外）", "メモ.txt"),
    ("情報", "案件A", "指摘0件", "UD消化_A.xlsx"),
    ("情報", "案件A", "指摘0件", "試験結果報告書_A.xlsx"),
]

# 件数: (全行, 自動, 手動, 除外, 要確認)
EXPECTED_TOTALS = {"全行": 63, "自動": 60, "手動": 0, "除外": 1, "要確認": 2}


@dataclass
class DeliverCase:
    """deliver（check で止まったあと overwrite）を実行したあとの期待結果。"""
    id: str
    project: str
    rel: str            # プロジェクトルートからの相対パス
    setup: str
    copy: str           # 期待する「納品コピー」
    at_dest: bool       # 納品先にあるか
    backup: bytes | None = None  # 退避されているべき中身（None なら退避なし）

    def expectation(self) -> str:
        parts = [f"納品コピー={self.copy}", "納品先にあり" if self.at_dest else "納品先になし"]
        parts.append("退避あり（上書き前の内容）" if self.backup is not None else "退避なし")
        return "、".join(parts)


DELIVER_CASES: list[DeliverCase] = [
    DeliverCase("D-1", "案件A", "030.調査/調査検討書_A.xlsx",
                "納品先に原本と同じ内容のファイルを置いた（コピー済みとみなす）", "済", True),
    DeliverCase("D-2", "案件A", "040.設計/成果物/外部Rv/040_レビューチェックリスト_設計(社外)(1回目)_A.xlsx",
                "納品先に原本と内容の違うファイルを置いた（check は止まり、overwrite で退避して上書き）", "済", True,
                backup=b"old version"),
    DeliverCase("D-3", "案件B", "030.調査/成果物/内部Rv/20261101/030_レビューチェックリスト_調査(社内)(1回目)_B.xlsx",
                "日付フォルダの中のチェックリスト（成果物/ の直下へ平らに置く）", "済", True),
    DeliverCase("D-4", "案件A", "080.SD作成/成果物/内部Rv/080_レビューチェックリスト_SD仕様書（社内）（2回目）_A.xlsx",
                "全角括弧のチェックリスト", "済", True),
    DeliverCase("D-5", "案件A", "040.設計/機能設計書_A.xlsx",
                "自社捺印・取引先捺印が未のドキュメント（前の工程が未）", "未", False),
    DeliverCase("D-6", "案件A", "040.設計/成果物/内部Rv/レビュー記録表_設計_A.xlsx",
                "対象が「要確認」の行", "未", False),
    DeliverCase("D-7", "案件A", "040.設計/メモ.txt", "対象が「除外」の行", "未", False),
]

# deliver のあと「納品コピー」が済になる行の数（捺印の要らない自動の行: 調査検討書2件・チェックリスト23件）
EXPECTED_COPIED = 25


@dataclass
class Result:
    id: str
    setup: str
    expected: str
    actual: str
    ok: bool


def _read_issues(path: Path) -> list[tuple[str, str, str, str, str]]:
    ws = load_workbook(path)["警告"]
    return [tuple("" if v is None else str(v) for v in row[:5]) for row in ws.iter_rows(min_row=2, values_only=True)]


def _verify_issues(issues) -> list[Result]:
    """警告シート全体: 期待どおりのものが出ていて、余計なものが出ていないか。"""
    results: list[Result] = []
    remaining = list(issues)
    for n, (lv, proj, cat, text) in enumerate(EXPECTED_ISSUES, start=1):
        hit = next((i for i in remaining if i[0] == lv and i[1] == proj and i[2] == cat and text in i[3]), None)
        if hit:
            remaining.remove(hit)
        results.append(Result(f"W-{n}", "警告シート", f"{lv} {proj}「{cat}」（{text}）",
                              "あり" if hit else "なし", hit is not None))
    for i in remaining:
        results.append(Result("W-余分", "警告シート", "（出ないはず）", f"{i[0]} {i[1]}「{i[2]}」{i[3]}", False))
    return results


def verify_deliver(demo_dir: Path) -> list[Result]:
    """deliver（check で止まったあと overwrite）を実行したあとの状態を照合する。"""
    cfg = load_config(demo_dir / "config.toml")
    delivery = load_delivery(demo_dir / "delivery.toml", cfg)
    records, _ = read_manifest(delivery.manifest_path)
    by_key = {(r.project, r.rel): r for r in records}
    results: list[Result] = []
    for c in DELIVER_CASES:
        r = by_key.get((c.project, c.rel))
        if r is None:
            results.append(Result(c.id, c.setup, c.expectation(), "管理表に載らない", False))
            continue
        name = c.rel.rsplit("/", 1)[-1]
        at_dest = any((delivery.dest_root / c.project).rglob(name)) if delivery.dest_root else False
        backup = Path(r.backup) if r.backup else None
        got_backup = backup.read_bytes() if backup and backup.is_file() else None
        ok = (r.st_copy == c.copy and at_dest == c.at_dest and got_backup == c.backup
              and (r.st_copy != "済" or bool(r.copy_hash)))
        actual = [f"No.{r.no} 納品コピー={r.st_copy}", "納品先にあり" if at_dest else "納品先になし",
                  "退避あり" + ("（上書き前の内容）" if got_backup == c.backup else "（内容が違う）")
                  if got_backup is not None else "退避なし"]
        if r.st_copy == "済":
            actual.append("ハッシュあり" if r.copy_hash else "ハッシュなし")
        results.append(Result(c.id, c.setup, c.expectation(), "、".join(actual), ok))
    copied = sum(1 for r in records if r.st_copy == "済")
    results.append(Result("T-2", "納品コピーが済の行", f"{EXPECTED_COPIED} 件", f"{copied} 件", copied == EXPECTED_COPIED))
    results += _verify_issues(_read_issues(delivery.manifest_path))
    return results


def verify(demo_dir: Path) -> list[Result]:
    dpath = demo_dir / "delivery.toml"
    cfg = load_config(demo_dir / "config.toml")
    delivery = load_delivery(dpath, cfg)
    records, _ = read_manifest(delivery.manifest_path)
    by_key = {(r.project, r.rel): r for r in records}
    issues = _read_issues(delivery.manifest_path)
    results: list[Result] = []

    for c in CASES:
        r = by_key.get((c.project, c.rel))
        problems, actual = [], []
        if not c.present:
            actual.append("管理表に載らない" if r is None else f"管理表に載った(No.{r.no})")
            if r is not None:
                problems.append("載ってはいけない")
        elif r is None:
            actual.append("管理表に載らない")
            problems.append("行がない")
        else:
            actual.append(f"No.{r.no} 対象={r.target}")
            if c.target and r.target != c.target:
                problems.append("対象")
            got = (r.st_share, r.st_self, r.st_send, r.st_partner, r.st_copy)
            if c.statuses:
                actual.append("ステータス=" + "/".join(got))
                if got != c.statuses:
                    problems.append("ステータス")
            if c.dest_rel:
                actual.append(f"納品先={r.dest_rel}")
                if r.dest_rel != c.dest_rel:
                    problems.append("納品先")
            if c.rv:
                got_rv = (r.rv_state, r.rv_progress, r.rv_approval_date)
                actual.append(f"状況={got_rv[0]}、進捗={got_rv[1]}、承認印の日付={got_rv[2] or '（空）'}")
                if got_rv != c.rv:
                    problems.append("記録表")
        for lv, cat in c.issues:
            # 行に紐づく警告は対象No、行がないもの（欠け）は内容の文字列で探す
            if r is not None:
                hit = any(i[0] == lv and i[2] == cat and str(r.no) in [x.strip() for x in i[4].split(",")]
                          for i in issues)
            else:
                name = c.rel.rsplit("/", 1)[-1].split("_")[0]
                hit = any(i[0] == lv and i[2] == cat and i[1] == c.project and name in i[3] for i in issues)
            actual.append(f"{lv}「{cat}」{'あり' if hit else 'なし'}")
            if not hit:
                problems.append(cat)
        results.append(Result(c.id, c.setup, c.expectation(), "、".join(actual), not problems))

    results += _verify_issues(issues)

    counts = {"全行": len(records)}
    for t in ("自動", "手動", "除外", "要確認"):
        counts[t] = sum(1 for r in records if r.target == t)
    exp = "、".join(f"{k}{v}" for k, v in EXPECTED_TOTALS.items())
    act = "、".join(f"{k}{v}" for k, v in counts.items())
    results.append(Result("T-1", "件数", exp, act, counts == EXPECTED_TOTALS))
    return results


def to_markdown(results: list[Result] | None = None) -> str:
    lines = [
        "# デモデータ 正誤表",
        "",
        "`python -m tests.dummy demo` で作るダミー3案件に仕込んだ内容と、期待する判定結果の一覧。",
        "実際の出力との突き合わせは `python -m tests.verify_demo` で行う（各行に ○/× が出る）。",
        "",
        "> このファイルは `python -m tests.verify_demo demo --markdown docs/デモ正誤表.md` で生成している。",
        "> 内容を変えるときは `tests/verify_demo.py` の CASES と `tests/dummy.py` を直してから再生成する。",
        "",
        "ステータスの並び: 指摘展開 / 自社捺印 / 取引先送付 / 取引先捺印 / 納品コピー",
        "",
        "## ファイルごとの期待結果",
        "",
        "| ID | 区分 | 案件 | ファイル（相対パス） | 仕込んだ内容 | 期待結果 |",
        "|---|---|---|---|---|---|",
    ]
    for c in CASES:
        lines.append(f"| {c.id} | {c.group} | {c.project} | `{c.rel}` | {c.setup} | {c.expectation()} |")
    lines += ["", "## 警告シートに出るべきもの（これ以外は出ない）", "",
              "| ID | 区分 | 案件 | 種類 | 内容に含まれる文字列 |", "|---|---|---|---|---|"]
    for n, (lv, proj, cat, text) in enumerate(EXPECTED_ISSUES, start=1):
        lines.append(f"| W-{n} | {lv} | {proj} | {cat} | {text} |")
    lines += ["", "## 件数", "", "| 全行 | 自動 | 手動 | 除外 | 要確認 |", "|---|---|---|---|---|",
              "| " + " | ".join(str(EXPECTED_TOTALS[k]) for k in ("全行", "自動", "手動", "除外", "要確認")) + " |"]
    lines += ["", "## 納品コピーのあとの期待結果", "",
              "collect のあと `deliver --mode check`（内容の違うファイルがあるので止まる）→ `deliver --mode overwrite` を実行し、"
              "`python -m tests.verify_demo --stage deliver` で照合する。",
              "納品先 `demo/dest/` には、collect の前から次のファイルを置いてある: 原本と同じ内容（D-1）、"
              "原本と内容が違う（D-2）、管理表にないファイル（W の「管理表にないファイル」）、Excel の一時ファイル（無視される）。",
              "警告シートは collect のあとと同じになる。",
              "", "| ID | 案件 | ファイル（相対パス） | 仕込んだ内容 | 期待結果 |", "|---|---|---|---|---|"]
    for c in DELIVER_CASES:
        lines.append(f"| {c.id} | {c.project} | `{c.rel}` | {c.setup} | {c.expectation()} |")
    lines += ["", f"納品コピーが済の行: {EXPECTED_COPIED} 件"]
    if results is not None:
        ok = sum(r.ok for r in results)
        lines += ["", f"## 照合結果（{ok}/{len(results)} 件 一致）", "", "| ID | 判定 | 実際の出力 |", "|---|---|---|"]
        for r in results:
            lines.append(f"| {r.id} | {'○' if r.ok else '×'} | {r.actual} |")
    return "\n".join(lines) + "\n"


def main(argv: list[str] | None = None) -> int:
    ap = argparse.ArgumentParser(description="デモデータの正誤表と実際の出力を突き合わせる")
    ap.add_argument("demo", nargs="?", default="demo", help="デモフォルダ（既定: demo）")
    ap.add_argument("--markdown", help="正誤表を Markdown で書き出すパス（照合結果は含めない）")
    ap.add_argument("--stage", choices=("collect", "deliver"), default="collect",
                    help="collect: collect のあとを照合（既定）/ deliver: deliver（overwrite）のあとを照合")
    args = ap.parse_args(argv)
    demo = Path(args.demo).resolve()

    if args.markdown:
        Path(args.markdown).parent.mkdir(parents=True, exist_ok=True)
        Path(args.markdown).write_text(to_markdown(), encoding="utf-8")
        print(f"正誤表を書き出しました: {args.markdown}")
        return 0

    results = verify_deliver(demo) if args.stage == "deliver" else verify(demo)
    width = max(len(r.id) for r in results)
    for r in results:
        print(f"{'○' if r.ok else '×'} {r.id:<{width}}  {r.setup}")
        if not r.ok:
            print(f"    期待: {r.expected}")
        print(f"    実際: {r.actual}")
    ng = [r for r in results if not r.ok]
    print()
    print(f"照合結果: {len(results) - len(ng)}/{len(results)} 件 一致" + ("" if not ng else f"、{len(ng)} 件 不一致"))
    return 1 if ng else 0


if __name__ == "__main__":
    sys.exit(main())
