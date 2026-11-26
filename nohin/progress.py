"""レビュー記録表の中身を読み、管理表の「記録表の状況」「指摘の進捗」「承認印の日付」を埋める。

前半はステータスへの自動反映はしない（印影の検出が必要なため後半で行う）。
"""

from __future__ import annotations

from pathlib import Path

from .config import Config
from .model import (KIND_RV, LEVEL_INFO, LEVEL_WARN, RV_ACTION_DONE, RV_IN_PROGRESS, RV_NO_FINDING, RV_NO_TABLE,
                    RV_UNREAD, TARGET_EXCLUDE, Issue, Record)
from .review_table import ReviewTableResult, range_label, read_review_table

MAX_LIST = 10  # 警告の内容に並べる行の最大数


def _rows_text(rows, describe=lambda r: r.label) -> str:
    items = [describe(r) for r in rows[:MAX_LIST]]
    if len(rows) > MAX_LIST:
        items.append(f"ほか{len(rows) - MAX_LIST}件")
    return "、".join(items)


def _is_editing(path: Path) -> bool:
    return any((path.parent / n).exists() for n in (f"~${path.name}", f"~${path.name[2:]}"))


def _clear(r: Record) -> None:
    r.rv_state = r.rv_progress = r.rv_approval_date = r.rv_sheet = r.rv_header_row = ""


def update_review_progress(records: list[Record], cfg: Config, now: str) -> list[Issue]:
    issues: list[Issue] = []
    partner_sides = cfg.review_record.partner_stamp_sides
    for r in records:
        if r.kind != KIND_RV or r.target == TARGET_EXCLUDE:
            _clear(r)
            continue
        path = Path(r.src)
        if not path.exists():  # 「元ファイルなし」は checks が出す
            _clear(r)
            r.rv_state = RV_UNREAD
            continue
        r.rv_read_at = now
        try:
            res = read_review_table(path, cfg.review_table)
            issues += _apply(r, res, internal=r.side not in partner_sides)
        except Exception as e:  # 1ファイルの想定外の問題で、全体の処理を止めない
            _clear(r)
            r.rv_state = RV_UNREAD
            issues.append(Issue(LEVEL_WARN, r.project, "記録表読取不可",
                                f"読み取り中にエラーが起きました（{type(e).__name__}: {e}）: {r.filename}", (r.id,)))
        if _is_editing(path):
            issues.append(Issue(LEVEL_INFO, r.project, "編集中",
                                f"Excel で開かれている可能性があります（保存前の内容を読んでいるかもしれません）: {r.filename}",
                                (r.id,)))
    return issues


def _apply(r: Record, res: ReviewTableResult, internal: bool) -> list[Issue]:
    ids = (r.id,)
    out: list[Issue] = []
    _clear(r)
    if not res.readable:
        r.rv_state = RV_UNREAD
        return [Issue(LEVEL_WARN, r.project, "記録表読取不可", f"{res.error}: {r.filename}", ids)]
    if not res.has_table:
        r.rv_state = RV_NO_TABLE
        return [Issue(LEVEL_WARN, r.project, "指摘表なし",
                      f"指摘の表の見出し（指摘事項・処置欄・日付・捺印(自社)）がそろう行が見つかりません: {r.filename}", ids)]

    r.rv_sheet, r.rv_header_row = res.sheet, str(res.header_row)
    findings, required = res.findings, res.required
    n_na = len(findings) - len(required)
    r.rv_progress = (f"指摘{len(findings)}" + (f"（捺印不要{n_na}）" if n_na else "")
                     + f"・処置{sum(1 for x in findings if x.action)}・日付{sum(1 for x in findings if x.date)}")
    if internal and res.latest_date:
        r.rv_approval_date = res.latest_date.isoformat()

    if not findings:
        r.rv_state = RV_NO_FINDING
        out.append(Issue(LEVEL_INFO, r.project, "指摘0件", f"指摘の表に指摘がありません: {r.filename}", ids))
    else:
        r.rv_state = RV_ACTION_DONE if res.action_done else RV_IN_PROGRESS

    missing = res.missing_dates()
    if missing:
        def why(x):
            return f"{x.label}（空）" if not x.date_raw.strip() else f"{x.label}（「{x.date_raw}」を日付として読めません）"
        out.append(Issue(LEVEL_WARN, r.project, "日付なし",
                         f"処置欄があるのに処置の日付がない行があります: {_rows_text(missing, why)}: {r.filename}", ids))
    if res.strays:
        out.append(Issue(LEVEL_WARN, r.project, "指摘欄が空の行に記入",
                         f"指摘事項が空なのに記入がある行があります: {_rows_text(res.strays)}: {r.filename}", ids))
    if not res.approvals:
        out.append(Issue(LEVEL_WARN, r.project, "承認欄なし",
                         f"指摘の表より上に承認欄（「承認」）が見つかりません: {r.filename}", ids))
    elif len(res.approvals) > 1:
        cells = "、".join(range_label(lab) for lab, _ in res.approvals)
        out.append(Issue(LEVEL_WARN, r.project, "承認欄なし",
                         f"承認欄（「承認」）が複数あり、どれか決められません（{cells}）: {r.filename}", ids))
    return out
