"""集計と完了判定。"""

from __future__ import annotations

from collections import Counter, defaultdict
from dataclasses import dataclass, field
from datetime import datetime

from .model import (KIND_RV, RV_ACTION_DONE, RV_IN_PROGRESS, RV_NO_FINDING, RV_NO_TABLE, RV_UNREAD, LEVEL_ERROR, LEVEL_INFO, LEVEL_WARN, STATUS_HEADERS, STATUS_KEYS, ST_DONE, ST_NA, ST_NG,
                    ST_TODO, TARGET_CHECK, TARGET_EXCLUDE, TARGET_MANUAL, TARGET_AUTO, Issue, Record)


@dataclass
class Summary:
    complete: bool
    reasons: list[str] = field(default_factory=list)
    rows: list[list] = field(default_factory=list)   # サマリーシートに書く表
    lines: list[str] = field(default_factory=list)   # 画面に出す要約


def apply_confirmations(issues: list[Issue], confirmed: dict[str, str]) -> None:
    for i in issues:
        if not i.confirmed:
            i.confirmed = confirmed.get(i.key, "")


def open_issues(issues: list[Issue]) -> list[Issue]:
    """完了を妨げる警告（エラー・警告で、確認済でないもの）。"""
    return [i for i in issues if i.level in (LEVEL_ERROR, LEVEL_WARN) and not i.confirmed]


def apply_row_warnings(records: list[Record], issues: list[Issue]) -> None:
    """管理表の「警告」列に、その行に関係する未確認の警告の種類を書く。"""
    labels: dict[str, list[str]] = defaultdict(list)
    for i in open_issues(issues):
        for rid in i.row_ids:
            if i.category not in labels[rid]:
                labels[rid].append(i.category)
    for r in records:
        r.warning = "・".join(labels.get(r.id, []))


def is_row_done(r: Record) -> bool:
    return all(getattr(r, k) in (ST_DONE, ST_NA) for k in STATUS_KEYS)


def summarize(delivery_name: str, records: list[Record], issues: list[Issue]) -> Summary:
    active = [r for r in records if r.target in (TARGET_AUTO, TARGET_MANUAL)]
    checking = [r for r in records if r.target == TARGET_CHECK]
    excluded = [r for r in records if r.target == TARGET_EXCLUDE]
    undone = [r for r in active if not is_row_done(r)]
    blocking = open_issues(issues)
    lvl = Counter(i.level for i in issues if not i.confirmed)

    reasons = []
    if checking:
        reasons.append(f"対象が「要確認」の行が {len(checking)} 件あります（自動・手動・除外のいずれかに変更してください）")
    if undone:
        reasons.append(f"未完了の行が {len(undone)} 件あります")
    if blocking:
        reasons.append(f"未確認のエラー・警告が {len(blocking)} 件あります（解消するか、警告シートで「確認済」にしてください）")
    if not active:
        reasons.append("対象の行がありません")
    complete = not reasons

    now = datetime.now().strftime("%Y-%m-%d %H:%M")
    rows: list[list] = [
        ["■ 納品管理表 サマリー"],
        ["納品名", delivery_name],
        ["更新日時", now],
        ["判定", "完了" if complete else "未完了"],
    ]
    for reason in reasons:
        rows.append(["", reason])
    rows += [
        [],
        ["■ 件数"],
        ["全行", len(records)],
        ["対象（自動・手動）", len(active)],
        ["要確認", len(checking)],
        ["除外", len(excluded)],
        ["完了した行", len(active) - len(undone)],
        ["エラー（未確認）", lvl.get(LEVEL_ERROR, 0)],
        ["警告（未確認）", lvl.get(LEVEL_WARN, 0)],
        ["情報", lvl.get(LEVEL_INFO, 0)],
        [],
        ["■ 工程別ステータス", ST_DONE, ST_TODO, ST_NG, ST_NA],
    ]
    for k in STATUS_KEYS:
        c = Counter(getattr(r, k) for r in active)
        rows.append([STATUS_HEADERS[k], c.get(ST_DONE, 0), c.get(ST_TODO, 0), c.get(ST_NG, 0), c.get(ST_NA, 0)])

    rv = [r for r in records if r.kind == KIND_RV and r.target != TARGET_EXCLUDE]
    if rv:
        rows += [[], ["■ レビュー記録表の状況", "件数"]]
        c = Counter(r.rv_state or "（未読取）" for r in rv)
        for state in (RV_ACTION_DONE, RV_IN_PROGRESS, RV_NO_FINDING, RV_NO_TABLE, RV_UNREAD, "（未読取）"):
            if c.get(state):
                rows.append([state, c[state]])

    rows += [[], ["■ 案件・工程別", "工程", "対象数", "完了", "未完了", "要確認"]]
    groups: dict[tuple, list[Record]] = defaultdict(list)
    for r in records:
        groups[(r.project, r.process)].append(r)
    for (proj, code), rs in sorted(groups.items()):
        act = [r for r in rs if r.target in (TARGET_AUTO, TARGET_MANUAL)]
        done = sum(1 for r in act if is_row_done(r))
        rows.append([proj, code, len(act), done, len(act) - done,
                     sum(1 for r in rs if r.target == TARGET_CHECK)])

    lines = [
        f"判定: {'完了' if complete else '未完了'}",
        *[f"  - {x}" for x in reasons],
        f"全 {len(records)} 行（対象 {len(active)} / 要確認 {len(checking)} / 除外 {len(excluded)}）",
        f"エラー {lvl.get(LEVEL_ERROR, 0)} 件 / 警告 {lvl.get(LEVEL_WARN, 0)} 件 / 情報 {lvl.get(LEVEL_INFO, 0)} 件（確認済を除く）",
    ]
    return Summary(complete=complete, reasons=reasons, rows=rows, lines=lines)
