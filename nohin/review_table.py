"""レビュー記録表の中身を読み取り専用で調べる。

- 指摘の表: 見出しの文字で見出し行を探し、データ行を分類する（行がずれても追従できる）。
- 承認欄: 見出し行より上で「承認」のラベルを探し、その直下のセルを承認印の枠とする。
- 画像・図形: xlsx の描画 XML から一覧を取り出す（前半は inspect での表示のみ）。

ファイルは保存しない。openpyxl はセルの値と結合セルの取得だけに使う。
"""

from __future__ import annotations

import posixpath
import re
import unicodedata
import warnings
import zipfile
from dataclasses import dataclass, field
from datetime import date
from pathlib import Path
from xml.etree import ElementTree as ET

from openpyxl import load_workbook
from openpyxl.utils import get_column_letter

from .config import ReviewTableConfig
from .dates import parse_date

# 行の分類
ROW_REQUIRED = "指摘（捺印必要）"
ROW_NOT_REQUIRED = "指摘（捺印不要）"
ROW_NO_FINDING_MARK = "指摘なしの印"
ROW_STRAY = "指摘欄が空の行に記入"
FINDING_KINDS = (ROW_REQUIRED, ROW_NOT_REQUIRED)

READABLE_EXTS = (".xlsx", ".xlsm")

CellRange = tuple[int, int, int, int]  # (min_row, min_col, max_row, max_col) 1 始まり


def norm(value) -> str:
    """比較用の文字列: NFKC 正規化して空白・改行を除く。"""
    if value is None:
        return ""
    return re.sub(r"\s+", "", unicodedata.normalize("NFKC", str(value)))


def range_label(rng: CellRange | None) -> str:
    if not rng:
        return ""
    r1, c1, r2, c2 = rng
    a = f"{get_column_letter(c1)}{r1}"
    return a if (r1, c1) == (r2, c2) else f"{a}:{get_column_letter(c2)}{r2}"


@dataclass
class TableRow:
    sheet: str
    row: int
    kind: str
    no: str = ""
    finding: str = ""
    action: str = ""
    date_raw: str = ""
    date: date | None = None
    self_mark: str = ""                 # 捺印(自社)セルの文字（「-」など）
    self_range: CellRange | None = None  # 押印セルの範囲（結合を考慮）
    partner_mark: str = ""
    partner_range: CellRange | None = None

    @property
    def label(self) -> str:
        return f"No.{self.no}" if self.no else f"{self.row}行目"


@dataclass
class ReviewTableResult:
    path: Path
    error: str = ""                     # 読めなかった理由（空なら読めた）
    sheet: str = ""
    header_row: int | None = None
    data_start: int | None = None
    columns: dict[str, int] = field(default_factory=dict)
    approvals: list[tuple[CellRange, CellRange]] = field(default_factory=list)  # (ラベル, 枠)
    rows: list[TableRow] = field(default_factory=list)
    sheets: list[str] = field(default_factory=list)

    @property
    def readable(self) -> bool:
        return not self.error

    @property
    def has_table(self) -> bool:
        return self.header_row is not None

    @property
    def findings(self) -> list[TableRow]:
        return [r for r in self.rows if r.kind in FINDING_KINDS]

    @property
    def required(self) -> list[TableRow]:
        return [r for r in self.rows if r.kind == ROW_REQUIRED]

    @property
    def strays(self) -> list[TableRow]:
        return [r for r in self.rows if r.kind == ROW_STRAY]

    @property
    def latest_date(self) -> date | None:
        ds = [r.date for r in self.findings if r.date]
        return max(ds) if ds else None

    def missing_dates(self) -> list[TableRow]:
        """捺印が必要な指摘のうち、処置があるのに日付が空・読めないもの。"""
        return [r for r in self.required if r.action and r.date is None]

    @property
    def action_done(self) -> bool:
        """捺印が必要な指摘すべてに処置と読める日付がある。"""
        return all(r.action and r.date for r in self.required)


class _Merged:
    def __init__(self, ws):
        self.ranges = [(m.min_row, m.min_col, m.max_row, m.max_col) for m in ws.merged_cells.ranges]

    def of(self, row: int, col: int) -> CellRange:
        for r1, c1, r2, c2 in self.ranges:
            if r1 <= row <= r2 and c1 <= col <= c2:
                return (r1, c1, r2, c2)
        return (row, col, row, col)


def _find_header(rows: list[tuple], cfg: ReviewTableConfig) -> tuple[int, dict[str, int]] | None:
    """見出し行（1 始まりの行番号）と列番号を返す。"""
    labels = {k: {norm(x) for x in getattr(cfg, k)} for k in
              ("finding_labels", "action_labels", "date_labels", "self_stamp_labels",
               "partner_stamp_labels", "no_labels")}
    for i, values in enumerate(rows, start=1):
        cells = [(c, norm(v)) for c, v in enumerate(values, start=1) if v is not None]

        def cols(key):
            return [c for c, t in cells if t in labels[key]]

        finding, action = cols("finding_labels"), cols("action_labels")
        if not finding or not action:
            continue
        a = action[0]
        right = lambda key: sorted(c for c in cols(key) if c > a)  # noqa: E731
        dates, stamps = right("date_labels"), right("self_stamp_labels")
        if not dates or not stamps:
            continue
        columns = {"finding": finding[0], "action": a, "date": dates[0], "self_stamp": stamps[0]}
        partner = cols("partner_stamp_labels")
        if partner:
            columns["partner_stamp"] = partner[0]
        no = [c for c in cols("no_labels") if c < a]
        if no:
            columns["no"] = no[0]
        return i, columns
    return None


def read_review_table(path: str | Path, cfg: ReviewTableConfig) -> ReviewTableResult:
    path = Path(path)
    res = ReviewTableResult(path=path)
    if path.suffix.lower() not in READABLE_EXTS:
        res.error = f"{path.suffix or '拡張子なし'} は読み取れません（.xlsx / .xlsm のみ）"
        return res
    try:
        with warnings.catch_warnings():
            warnings.simplefilter("ignore")
            wb = load_workbook(path, data_only=True)
    except Exception as e:  # 壊れたファイル・パスワード付きなど
        res.error = f"開けません（壊れている・パスワード付き・Excel 以外の可能性）: {type(e).__name__}"
        return res

    try:
        res.sheets = list(wb.sheetnames)
        for ws in wb.worksheets:
            if not hasattr(ws, "iter_rows") or not hasattr(ws, "merged_cells"):
                continue  # グラフシートなど、セルのないシートは飛ばす
            head = list(ws.iter_rows(min_row=1, max_row=cfg.header_search_rows, values_only=True))
            found = _find_header(head, cfg)
            if not found:
                continue
            res.sheet = ws.title
            res.header_row, res.columns = found
            merged = _Merged(ws)
            start = max(merged.of(res.header_row, res.columns[k])[2] for k in
                        ("finding", "action", "date", "self_stamp")) + 1
            res.data_start = start
            res.rows = _read_rows(ws, merged, start, res.columns, cfg)
            res.approvals = _find_approvals(head[:res.header_row - 1], merged, cfg)
            break
    finally:
        wb.close()
    return res


def _read_rows(ws, merged: _Merged, start: int, cols: dict[str, int], cfg: ReviewTableConfig) -> list[TableRow]:
    na = {norm(x) for x in cfg.not_required_marks}
    no_finding = {norm(x) for x in cfg.no_finding_marks}
    out: list[TableRow] = []
    blank = 0

    def at(values, key):
        c = cols.get(key)
        return values[c - 1] if c and c - 1 < len(values) else None

    for r, values in enumerate(ws.iter_rows(min_row=start, values_only=True), start=start):
        finding, action = norm(at(values, "finding")), norm(at(values, "action"))
        draw = at(values, "date")
        dtext = norm(draw)
        smark, pmark = norm(at(values, "self_stamp")), norm(at(values, "partner_stamp"))
        stamp_blank = smark in ("", *na) and pmark in ("", *na)

        if not finding and not action and not dtext and stamp_blank:
            blank += 1
            if blank >= cfg.blank_stop_rows:
                break
            continue
        blank = 0

        if finding in no_finding:
            kind = ROW_NO_FINDING_MARK
        elif not finding:
            kind = ROW_STRAY
        elif smark in na:
            kind = ROW_NOT_REQUIRED
        else:
            kind = ROW_REQUIRED
        no_val = at(values, "no")
        out.append(TableRow(
            sheet=ws.title, row=r, kind=kind,
            no=norm(no_val) if no_val is not None else "",
            finding=finding, action=action,
            date_raw="" if draw is None else str(draw),
            date=parse_date(draw),
            self_mark=smark, self_range=merged.of(r, cols["self_stamp"]),
            partner_mark=pmark,
            partner_range=merged.of(r, cols["partner_stamp"]) if "partner_stamp" in cols else None,
        ))
    return out


def _find_approvals(rows: list[tuple], merged: _Merged, cfg: ReviewTableConfig) -> list[tuple[CellRange, CellRange]]:
    labels = {norm(x) for x in cfg.approval_labels}
    found = []
    for r, values in enumerate(rows, start=1):
        for c, v in enumerate(values, start=1):
            if v is not None and norm(v) in labels:
                lab = merged.of(r, c)
                if lab[:2] != (r, c):
                    continue  # 結合セルの左上以外は飛ばす
                box = merged.of(lab[2] + 1, lab[1])
                found.append((lab, box))
    return found


# ---------------------------------------------------------------- 画像・図形の一覧（inspect 用）

NS_MAIN = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
NS_REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
NS_PKG_REL = "http://schemas.openxmlformats.org/package/2006/relationships"
ANCHORS = ("twoCellAnchor", "oneCellAnchor", "absoluteAnchor")
OBJECT_KINDS = {"pic": "画像", "sp": "図形", "grpSp": "グループ", "graphicFrame": "グラフ等", "cxnSp": "線"}


@dataclass
class DrawingObject:
    sheet: str
    kind: str
    name: str
    anchor: str
    from_cell: tuple[int, int] | None  # (row, col) 1 始まり
    to_cell: tuple[int, int] | None

    def cells(self) -> str:
        def lab(rc):
            return f"{get_column_letter(rc[1])}{rc[0]}" if rc else "?"
        if self.from_cell is None:
            return "（セルに固定されていない）"
        return lab(self.from_cell) if not self.to_cell else f"{lab(self.from_cell)}〜{lab(self.to_cell)}"


def _local(tag: str) -> str:
    return tag.rsplit("}", 1)[-1]


def _rels(zf: zipfile.ZipFile, part: str) -> dict[str, tuple[str, str]]:
    base_dir, name = posixpath.split(part)
    rel_path = posixpath.join(base_dir, "_rels", name + ".rels")
    if rel_path not in zf.namelist():
        return {}
    root = ET.fromstring(zf.read(rel_path))
    out = {}
    for rel in root:
        target = rel.get("Target", "")
        resolved = target.lstrip("/") if target.startswith("/") else posixpath.normpath(posixpath.join(base_dir, target))
        out[rel.get("Id")] = (rel.get("Type", ""), resolved)
    return out


def _cell(el) -> tuple[int, int] | None:
    if el is None:
        return None
    vals = {_local(c.tag): c.text for c in el}
    try:
        return int(vals["row"]) + 1, int(vals["col"]) + 1
    except (KeyError, TypeError, ValueError):
        return None


def list_drawings(path: str | Path) -> list[DrawingObject]:
    """ファイル内の画像・図形を、シートごとにアンカーのセルと一緒に返す。"""
    out: list[DrawingObject] = []
    with zipfile.ZipFile(path) as zf:
        wb_rels = _rels(zf, "xl/workbook.xml")
        wb = ET.fromstring(zf.read("xl/workbook.xml"))
        for sh in wb.iter(f"{{{NS_MAIN}}}sheet"):
            rid = sh.get(f"{{{NS_REL}}}id")
            if rid not in wb_rels:
                continue
            sheet_part = wb_rels[rid][1]
            for typ, target in _rels(zf, sheet_part).values():
                if not typ.endswith("/drawing") or target not in zf.namelist():
                    continue
                root = ET.fromstring(zf.read(target))
                for anchor in root.iter():
                    aname = _local(anchor.tag)
                    if aname not in ANCHORS:
                        continue
                    obj = next((c for c in anchor if _local(c.tag) in OBJECT_KINDS), None)
                    kind = OBJECT_KINDS.get(_local(obj.tag), "不明") if obj is not None else "不明"
                    cnv = next((e for e in (obj.iter() if obj is not None else []) if _local(e.tag) == "cNvPr"), None)
                    frm = next((c for c in anchor if _local(c.tag) == "from"), None)
                    to = next((c for c in anchor if _local(c.tag) == "to"), None)
                    out.append(DrawingObject(sheet=sh.get("name", ""), kind=kind,
                                             name=cnv.get("name", "") if cnv is not None else "",
                                             anchor=aname, from_cell=_cell(frm), to_cell=_cell(to)))
    return out
