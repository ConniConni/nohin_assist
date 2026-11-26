"""管理表（xlsx）の作成・読み込み・更新。

- 管理表シートは、人が並べ替えや書式変更をしても壊れないよう、ID 列で行を探して値だけを書き換える。
- 見出しにない列（人が追加した列）には触らない。
- サマリー・警告シートは毎回作り直す。警告シートの「確認済」は引き継ぐ。
"""

from __future__ import annotations

import shutil
from datetime import datetime
from pathlib import Path

from openpyxl import Workbook, load_workbook
from openpyxl.formatting.rule import CellIsRule
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.datavalidation import DataValidation

from .model import (ATTR_TO_HEADER, COLUMNS, HEADER_TO_ATTR, LEVEL_ERROR, LEVEL_WARN, STATUS_KEYS, STATUSES,
                    TARGETS, Issue, Record)

SHEET_MAIN = "管理表"
SHEET_SUMMARY = "サマリー"
SHEET_ISSUES = "警告"
VALIDATION_ROWS = 5000
REQUIRED_HEADERS = ("ID", "ファイル名", "対象") + tuple(ATTR_TO_HEADER[k] for k in STATUS_KEYS)

FILL_HEADER = PatternFill("solid", fgColor="E7E6E6")
FILL_DONE = PatternFill("solid", fgColor="C6EFCE")
FILL_NG = PatternFill("solid", fgColor="FFC7CE")
FILL_CHECK = PatternFill("solid", fgColor="FFEB9C")
FONT_NA = Font(color="A6A6A6")
FONT_BOLD = Font(bold=True)


class ManifestError(Exception):
    """管理表を読み書きできない。"""


# ---------------------------------------------------------------- 排他・バックアップ

def ensure_not_open(path: Path) -> None:
    """Excel で開いているときにできる ~$ ファイルがあれば止める。"""
    candidates = [path.parent / f"~${path.name}", path.parent / f"~${path.name[2:]}"]
    if any(c.exists() for c in candidates):
        raise ManifestError(f"管理表が Excel で開かれています。閉じてから再実行してください: {path}")
    if path.exists():
        try:
            with path.open("r+b"):
                pass
        except PermissionError as e:
            raise ManifestError(f"管理表に書き込めません（開いている可能性があります）: {path}") from e


def backup_manifest(path: Path, work_dir: Path, label: str) -> Path | None:
    if not path.exists():
        return None
    ts = datetime.now().strftime("%Y%m%d-%H%M%S")
    dest = work_dir / "_backup" / f"{ts}_{label}" / path.name
    dest.parent.mkdir(parents=True, exist_ok=True)
    shutil.copy2(path, dest)
    return dest


# ---------------------------------------------------------------- 読み込み

def _cell_str(v) -> str:
    if v is None:
        return ""
    if isinstance(v, datetime):
        return v.strftime("%Y-%m-%d %H:%M")
    if isinstance(v, float) and v.is_integer():
        v = int(v)
    return str(v).strip()


def _header_map(ws) -> dict[str, int]:
    hm = {}
    for cell in ws[1]:
        if cell.value is not None:
            hm[str(cell.value).strip()] = cell.column
    return hm


def read_manifest(path: Path) -> tuple[list[Record], dict[str, str]]:
    """管理表を読む。戻り値: (行の一覧, 警告キー → 確認済マーク)。"""
    if not path.exists():
        raise ManifestError(f"管理表がありません。先に collect を実行してください: {path}")
    try:
        wb = load_workbook(path)
    except Exception as e:  # 壊れたファイル・別形式
        raise ManifestError(f"管理表を開けません: {path} ({e})") from e
    if SHEET_MAIN not in wb.sheetnames:
        raise ManifestError(f"管理表に「{SHEET_MAIN}」シートがありません: {path}")
    ws = wb[SHEET_MAIN]
    hm = _header_map(ws)
    missing = [h for h in REQUIRED_HEADERS if h not in hm]
    if missing:
        raise ManifestError(f"管理表の見出しが足りません（列を消したり名前を変えたりしていませんか）: {missing}")

    records = []
    seen = set()
    for row in range(2, ws.max_row + 1):
        rid = _cell_str(ws.cell(row, hm["ID"]).value)
        if not rid:
            continue
        if rid in seen:
            raise ManifestError(f"管理表の ID が重複しています（行をコピーしていませんか）: {rid}")
        seen.add(rid)
        rec = Record()
        for header, col in hm.items():
            attr = HEADER_TO_ATTR.get(header)
            if not attr:
                continue
            v = _cell_str(ws.cell(row, col).value)
            if attr in ("no", "round"):
                setattr(rec, attr, int(v) if v.isdigit() else None)
            else:
                setattr(rec, attr, v)
        records.append(rec)

    confirmed = {}
    if SHEET_ISSUES in wb.sheetnames:
        wi = wb[SHEET_ISSUES]
        ih = _header_map(wi)
        if "キー" in ih and "確認済" in ih:
            for row in range(2, wi.max_row + 1):
                k = _cell_str(wi.cell(row, ih["キー"]).value)
                c = _cell_str(wi.cell(row, ih["確認済"]).value)
                if k and c:
                    confirmed[k] = c
    return records, confirmed


# ---------------------------------------------------------------- 書き込み

def _new_workbook() -> Workbook:
    wb = Workbook()
    ws = wb.active
    ws.title = SHEET_MAIN
    for i, (attr, header, width, detail) in enumerate(COLUMNS, start=1):
        c = ws.cell(1, i, header)
        c.fill, c.font = FILL_HEADER, FONT_BOLD
        c.alignment = Alignment(vertical="center")
        ws.column_dimensions[get_column_letter(i)].width = width
    last = get_column_letter(len(COLUMNS))
    filename_col = [a for a, *_ in COLUMNS].index("filename") + 1
    ws.freeze_panes = ws.cell(2, filename_col + 1)
    ws.auto_filter.ref = f"A1:{last}{VALIDATION_ROWS}"

    def col_of(attr):
        return get_column_letter([a for a, *_ in COLUMNS].index(attr) + 1)

    dv_target = DataValidation(type="list", formula1='"' + ",".join(TARGETS) + '"', allow_blank=False)
    dv_status = DataValidation(type="list", formula1='"' + ",".join(STATUSES) + '"', allow_blank=False)
    ws.add_data_validation(dv_target)
    ws.add_data_validation(dv_status)
    dv_target.add(f"{col_of('target')}2:{col_of('target')}{VALIDATION_ROWS}")
    st_range = f"{col_of(STATUS_KEYS[0])}2:{col_of(STATUS_KEYS[-1])}{VALIDATION_ROWS}"
    dv_status.add(st_range)

    cf = ws.conditional_formatting
    cf.add(st_range, CellIsRule(operator="equal", formula=['"済"'], fill=FILL_DONE))
    cf.add(st_range, CellIsRule(operator="equal", formula=['"NG"'], fill=FILL_NG))
    cf.add(st_range, CellIsRule(operator="equal", formula=['"不要"'], font=FONT_NA))
    tg_range = f"{col_of('target')}2:{col_of('target')}{VALIDATION_ROWS}"
    cf.add(tg_range, CellIsRule(operator="equal", formula=['"要確認"'], fill=FILL_CHECK))
    cf.add(tg_range, CellIsRule(operator="equal", formula=['"除外"'], font=FONT_NA))
    ws.column_dimensions[col_of("warning")].font = Font(color="C00000")

    detail_cols = [i for i, (_, _, _, d) in enumerate(COLUMNS, start=1) if d]
    if detail_cols:
        ws.column_dimensions.group(get_column_letter(detail_cols[0]), get_column_letter(detail_cols[-1]),
                                   hidden=True, outline_level=1)
    ws.sheet_properties.outlinePr.summaryRight = False
    return wb


def save_manifest(path: Path, records: list[Record], issues: list[Issue], summary_rows: list[list],
                  confirmed: dict[str, str] | None = None) -> None:
    """管理表を保存する（なければ新規作成）。records の順に No を振り直す。"""
    ensure_not_open(path)
    path.parent.mkdir(parents=True, exist_ok=True)
    wb = load_workbook(path) if path.exists() else _new_workbook()
    ws = wb[SHEET_MAIN]
    hm = _header_map(ws)
    for attr, header, *_ in COLUMNS:  # 人が消した列は末尾に足し直す
        if header not in hm:
            col = ws.max_column + 1
            ws.cell(1, col, header).font = FONT_BOLD
            hm[header] = col

    id_col = hm["ID"]
    row_of = {}
    for row in range(2, ws.max_row + 1):
        v = _cell_str(ws.cell(row, id_col).value)
        if v:
            row_of[v] = row
    next_row = max([1, *row_of.values()]) + 1

    for i, rec in enumerate(records, start=1):
        rec.no = i
        row = row_of.get(rec.id)
        if row is None:
            row, next_row = next_row, next_row + 1
        for attr, header, *_ in COLUMNS:
            v = getattr(rec, attr)
            ws.cell(row, hm[header], "" if v is None else v)

    _write_issues(wb, records, issues, confirmed or {})
    _write_summary(wb, summary_rows)
    wb.active = wb.sheetnames.index(SHEET_MAIN)
    tmp = path.with_name(path.stem + ".tmp.xlsx")
    wb.save(tmp)
    tmp.replace(path)


def _replace_sheet(wb, name: str, index: int):
    if name in wb.sheetnames:
        del wb[name]
    return wb.create_sheet(name, index)


def _write_issues(wb, records: list[Record], issues: list[Issue], confirmed: dict[str, str]) -> None:
    ws = _replace_sheet(wb, SHEET_ISSUES, len(wb.sheetnames))
    no_of = {r.id: r.no for r in records}
    headers = [("区分", 8), ("案件", 14), ("種類", 14), ("内容", 90), ("対象No", 12), ("確認済", 8), ("キー", 30)]
    for i, (h, w) in enumerate(headers, start=1):
        c = ws.cell(1, i, h)
        c.fill, c.font = FILL_HEADER, FONT_BOLD
        ws.column_dimensions[get_column_letter(i)].width = w
    for r, iss in enumerate(issues, start=2):
        nos = ", ".join(str(no_of[x]) for x in iss.row_ids if x in no_of)
        mark = iss.confirmed or confirmed.get(iss.key, "")
        values = [iss.level, iss.project, iss.category, iss.message, nos, mark, iss.key]
        for c, v in enumerate(values, start=1):
            ws.cell(r, c, v)
        if iss.level == LEVEL_ERROR:
            ws.cell(r, 1).fill = FILL_NG
        elif iss.level == LEVEL_WARN:
            ws.cell(r, 1).fill = FILL_CHECK
    ws.freeze_panes = "A2"
    ws.auto_filter.ref = f"A1:G{max(2, len(issues) + 1)}"
    ws.column_dimensions["G"].hidden = True
    dv = DataValidation(type="list", formula1='"確認済"', allow_blank=True)
    ws.add_data_validation(dv)
    dv.add(f"F2:F{max(2, len(issues) + 200)}")


def _write_summary(wb, rows: list[list]) -> None:
    ws = _replace_sheet(wb, SHEET_SUMMARY, 0)
    widths = [16, 14, 10, 10, 10, 10, 10]
    for i, w in enumerate(widths, start=1):
        ws.column_dimensions[get_column_letter(i)].width = w
    for r, row in enumerate(rows, start=1):
        for c, v in enumerate(row, start=1):
            cell = ws.cell(r, c, v)
            if isinstance(v, str) and v.startswith("■"):
                cell.font = FONT_BOLD
        if row and row[0] == "判定":
            ws.cell(r, 2).fill = FILL_DONE if row[1] == "完了" else FILL_CHECK
