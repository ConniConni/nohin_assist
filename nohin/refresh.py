"""collect・report・deliver・restore で共通の更新処理。

記録表の読み取り → 整合チェック → 納品先のチェック → 確認済の引き継ぎ → 集計。
"""

from __future__ import annotations

from pathlib import Path

from .checks import check_all, sort_issues
from .classify import enrich
from .destination import check_destination
from .model import Issue, Record
from .progress import update_review_progress
from .report import Summary, apply_confirmations, apply_row_warnings, summarize


def backup_stats(work_dir: Path) -> tuple[int, int]:
    """<work_dir>/_backup の下のファイル数と合計バイト数。"""
    root = work_dir / "_backup"
    if not root.is_dir():
        return 0, 0
    files = [p for p in root.rglob("*") if p.is_file()]
    return len(files), sum(p.stat().st_size for p in files)


def refresh(records: list[Record], cfg, delivery, extra: list[Issue], confirmed: dict[str, str],
            now: str) -> tuple[Summary, list[Issue]]:
    for r in records:
        enrich(r, cfg)
    issues = (extra + update_review_progress(records, cfg, now) + check_all(records, cfg, delivery)
              + check_destination(records, cfg, delivery))
    issues = sort_issues(issues)
    apply_confirmations(issues, confirmed)
    apply_row_warnings(records, issues)
    return summarize(delivery.name, records, issues, backup_stats(delivery.work_dir)), issues
