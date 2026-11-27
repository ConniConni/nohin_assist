"""コマンドライン入口。

    python -m nohin collect [-d delivery.toml] [-c config.toml] [--check-names] [--new]
    python -m nohin report  [-d delivery.toml] [-c config.toml]
"""

from __future__ import annotations

import argparse
import sys
from datetime import datetime
from pathlib import Path

from openpyxl.utils import get_column_letter

from . import __version__
from .checks import check_all, sort_issues
from .classify import build_record, enrich, in_scope
from .config import ConfigError, find_config, find_config_for_inspect, load_config, load_delivery
from .manifest import ManifestError, backup_manifest, ensure_not_open, read_manifest, save_manifest
from .model import KIND_CL, KIND_DOC, KIND_OTHER, KIND_RV, LEVEL_INFO, Issue, Record
from .progress import update_review_progress
from .report import apply_confirmations, apply_row_warnings, summarize
from .review_table import ROW_NOT_REQUIRED, list_drawings, range_label, read_review_table
from .scan import scan_project

KIND_ORDER = {KIND_DOC: 0, KIND_RV: 1, KIND_CL: 2, KIND_OTHER: 3}
MAX_PRINT = 30


def _load(args):
    dpath = Path(args.delivery).resolve()
    cfg = load_config(find_config(dpath, args.config))
    return cfg, load_delivery(dpath, cfg)


def _sort_key(project_order: dict[str, int]):
    def key(r: Record):
        return (project_order.get(r.project, 99), r.process, KIND_ORDER.get(r.kind, 9),
                r.side, r.round or 0, r.filename)
    return key


def _print_issues(issues: list[Issue]) -> None:
    shown = [i for i in issues if i.level != LEVEL_INFO and not i.confirmed]
    for i in shown[:MAX_PRINT]:
        print(f"  [{i.level}] {i.project} {i.category}: {i.message}")
    rest = len(shown) - MAX_PRINT
    if rest > 0:
        print(f"  …ほか {rest} 件（管理表の「警告」シートを参照）")
    info = sum(1 for i in issues if i.level == LEVEL_INFO)
    if info:
        print(f"  （情報 {info} 件は警告シートに記載）")


def cmd_collect(args) -> int:
    cfg, delivery = _load(args)
    now = datetime.now().strftime("%Y-%m-%d %H:%M")
    scanned_records: list[Record] = []
    scan_issues: list[Issue] = []
    for proj in delivery.projects:
        files, iss = scan_project(proj, cfg)
        scan_issues += iss
        built = (build_record(proj.name, f.rel, str(f.path), cfg, now) for f in files)
        scanned_records += [r for r in built if in_scope(r)]
    order = {p.name: i for i, p in enumerate(delivery.projects)}
    scanned_records.sort(key=_sort_key(order))

    if args.check_names:
        return _print_check_names(scanned_records, scan_issues, cfg, delivery)

    path = delivery.manifest_path
    ensure_not_open(path)
    confirmed: dict[str, str] = {}
    if path.exists() and not args.new:
        existing, confirmed = read_manifest(path)
        known = {r.id for r in existing}
        added = [r for r in scanned_records if r.id not in known]
        records = existing + added
        bk = backup_manifest(path, delivery.work_dir, "collect")
        print(f"既存の管理表に追記します（新規 {len(added)} 件）。更新前の管理表: {bk}")
    else:
        if path.exists():
            bk = backup_manifest(path, delivery.work_dir, "collect-new")
            path.unlink()
            print(f"管理表を作り直します。以前の管理表: {bk}")
        records = scanned_records

    summary, issues = _refresh(records, cfg, delivery, scan_issues, confirmed, now)
    save_manifest(path, records, issues, summary.rows, confirmed)

    print(f"管理表を保存しました: {path}")
    print("\n".join(summary.lines))
    _print_issues(issues)
    return 0


def _refresh(records: list[Record], cfg, delivery, extra: list[Issue], confirmed: dict[str, str], now: str):
    """記録表の読み取り → 整合チェック → 確認済の引き継ぎ → 集計。collect と report で共通。"""
    for r in records:
        enrich(r, cfg)
    issues = extra + update_review_progress(records, cfg, now) + check_all(records, cfg, delivery)
    issues = sort_issues(issues)
    apply_confirmations(issues, confirmed)
    apply_row_warnings(records, issues)
    return summarize(delivery.name, records, issues), issues


def _print_check_names(records: list[Record], scan_issues: list[Issue], cfg, delivery) -> int:
    print("判定結果（管理表は作成・変更しません）")
    current = None
    for r in records:
        if r.project != current:
            current = r.project
            print(f"\n=== {current} ===")
        detail = r.doc_name if r.kind == KIND_DOC else (f"{r.side} {r.round}回目" if r.side else "")
        print(f"  {r.process}\t{r.kind}\t{detail}\t{r.target}\t{r.rel}")
    issues = sort_issues(scan_issues + check_all(records, cfg, delivery))
    print()
    print(f"警告など: {len(issues)} 件")
    _print_issues(issues)
    return 0


def cmd_report(args) -> int:
    cfg, delivery = _load(args)
    path = delivery.manifest_path
    records, confirmed = read_manifest(path)
    now = datetime.now().strftime("%Y-%m-%d %H:%M")
    summary, issues = _refresh(records, cfg, delivery, [], confirmed, now)
    save_manifest(path, records, issues, summary.rows, confirmed)
    print(f"管理表のサマリー・警告を更新しました: {path}")
    print("\n".join(summary.lines))
    _print_issues(issues)
    return 0


def cmd_inspect(args) -> int:
    """レビュー記録表1つの読み取り結果を表示する（読み取り専用）。"""
    cfg = load_config(find_config_for_inspect(args.config))
    path = Path(args.file).resolve()
    if not path.exists():
        print(f"エラー: ファイルが見つかりません: {path}", file=sys.stderr)
        return 2
    rt = cfg.review_table
    res = read_review_table(path, rt)
    print(f"ファイル: {path}")
    print(f"設定: {cfg.path}")
    if not res.readable:
        print(f"読み取れません: {res.error}")
        return 1
    print(f"シート: {', '.join(res.sheets)}")
    if not res.has_table:
        print("指摘の表: 見つかりません")
        print(f"  探した見出し: 指摘事項={list(rt.finding_labels)} / 処置={list(rt.action_labels)} / "
              f"日付={list(rt.date_labels)} / 捺印(自社)={list(rt.self_stamp_labels)}")
    else:
        names = {"no": "No.", "finding": "指摘事項", "action": "処置欄", "date": "日付",
                 "self_stamp": "捺印(自社)", "partner_stamp": "捺印(取引先)"}
        cols = "、".join(f"{names[k]}={get_column_letter(c)}列" for k, c in res.columns.items())
        print(f"指摘の表: シート「{res.sheet}」の {res.header_row} 行目が見出し（データは {res.data_start} 行目から）")
        print(f"  列: {cols}")
        if not res.approvals:
            print("承認欄: 見つかりません（見出し行より上に「承認」がない）")
        for lab, box in res.approvals:
            print(f"承認欄: ラベル {range_label(lab)} → 承認印の枠 {range_label(box)}")
        print(f"行（{len(res.rows)} 件。空行は省略）:")
        for x in res.rows:
            date = x.date.isoformat() if x.date else (f"読めない「{x.date_raw}」" if x.date_raw.strip() else "なし")
            mark = "" if x.kind != ROW_NOT_REQUIRED else f"（捺印欄「{x.self_mark}」）"
            print(f"  {x.row}行目 {x.label:<8} {x.kind}{mark}  処置={'あり' if x.action else 'なし'}  日付={date}"
                  f"  自社印の枠={range_label(x.self_range)}"
                  + (f"  取引先印の枠={range_label(x.partner_range)}" if x.partner_range else ""))
        latest = res.latest_date.isoformat() if res.latest_date else "なし"
        print(f"最新の処置日（社内の承認印の日付）: {latest}")
        print(f"処置: {'完了' if res.action_done else '未完了'}"
              f"（指摘 {len(res.findings)} 件、うち捺印不要 {len(res.findings) - len(res.required)} 件）")
    try:
        drawings = list_drawings(path)
    except Exception as e:  # 描画の解析に失敗しても、ほかの結果は表示する
        print(f"画像・図形: 読み取れません（{type(e).__name__}）")
        return 0
    print(f"画像・図形: {len(drawings)} 件")
    for d in drawings:
        print(f"  シート「{d.sheet}」 {d.kind} 「{d.name}」 {d.cells()}（{d.anchor}）")
    return 0


def main(argv: list[str] | None = None) -> int:
    ap = argparse.ArgumentParser(prog="nohin", description=f"納品サポートツール v{__version__}")
    sub = ap.add_subparsers(dest="cmd", required=True)

    def common(p):
        p.add_argument("-d", "--delivery", default="delivery.toml", help="納品ごとの設定ファイル（既定: delivery.toml）")
        p.add_argument("-c", "--config", default=None, help="共通設定ファイル（既定: delivery.toml と同じ場所かツールのフォルダの config.toml）")

    p = sub.add_parser("collect", help="納品対象を洗い出して管理表を作成・追記する")
    common(p)
    p.add_argument("--check-names", action="store_true", help="判定結果を表示するだけ（管理表は作らない）")
    p.add_argument("--new", action="store_true", help="既存の管理表を退避して作り直す")
    p.set_defaults(func=cmd_collect)

    p = sub.add_parser("report", help="管理表を集計し、サマリーと警告を更新する")
    common(p)
    p.set_defaults(func=cmd_report)

    p = sub.add_parser("inspect", help="レビュー記録表1つの読み取り結果を表示する（読み取り専用）")
    p.add_argument("file", help="レビュー記録表（.xlsx）")
    p.add_argument("-c", "--config", default=None, help="共通設定ファイル（既定: ツールのフォルダの config.toml、なければ config.example.toml）")
    p.set_defaults(func=cmd_inspect)

    args = ap.parse_args(argv)
    try:
        return args.func(args)
    except (ConfigError, ManifestError) as e:
        print(f"エラー: {e}", file=sys.stderr)
        return 2
