"""コマンドライン入口。

    python -m nohin collect [-d delivery.toml] [-c config.toml] [--check-names] [--new]
    python -m nohin report  [-d delivery.toml] [-c config.toml]
"""

from __future__ import annotations

import argparse
import sys
from datetime import datetime
from pathlib import Path

from . import __version__
from .checks import check_all, sort_issues
from .classify import build_record, enrich
from .config import ConfigError, find_config, load_config, load_delivery
from .manifest import ManifestError, backup_manifest, ensure_not_open, read_manifest, save_manifest
from .model import KIND_CL, KIND_DOC, KIND_OTHER, KIND_RV, LEVEL_INFO, Issue, Record
from .report import apply_confirmations, apply_row_warnings, summarize
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
        scanned_records += [build_record(proj.name, f.rel, str(f.path), cfg, now) for f in files]
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

    for r in records:
        enrich(r, cfg)
    issues = sort_issues(scan_issues + check_all(records, cfg, delivery))
    apply_confirmations(issues, confirmed)
    apply_row_warnings(records, issues)
    summary = summarize(delivery.name, records, issues)
    save_manifest(path, records, issues, summary.rows, confirmed)

    print(f"管理表を保存しました: {path}")
    print("\n".join(summary.lines))
    _print_issues(issues)
    return 0


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
    for r in records:
        enrich(r, cfg)
    issues = sort_issues(check_all(records, cfg, delivery))
    apply_confirmations(issues, confirmed)
    apply_row_warnings(records, issues)
    summary = summarize(delivery.name, records, issues)
    save_manifest(path, records, issues, summary.rows, confirmed)
    print(f"管理表のサマリー・警告を更新しました: {path}")
    print("\n".join(summary.lines))
    _print_issues(issues)
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

    args = ap.parse_args(argv)
    try:
        return args.func(args)
    except (ConfigError, ManifestError) as e:
        print(f"エラー: {e}", file=sys.stderr)
        return 2
