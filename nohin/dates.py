"""セルの値を日付として解釈する（レビュー記録表の処置日、Step 4 の印影の日付で共通に使う）。"""

from __future__ import annotations

import re
import unicodedata
from datetime import date, datetime

_ERA_BASE = {"令和": 2018, "R": 2018, "平成": 1988, "H": 1988}
_TIME = r"(\d{1,2}:\d{2}(:\d{2})?)?"
_YMD = re.compile(r"(\d{4})[/\-.年](\d{1,2})[/\-.月](\d{1,2})日?" + _TIME)
_ERA = re.compile(r"(令和|平成|R|H)(\d{1,2}|元)[/\-.年](\d{1,2})[/\-.月](\d{1,2})日?" + _TIME, re.IGNORECASE)


def parse_date(value) -> date | None:
    """日付として読めれば date、読めなければ None（年がないもの・数値だけのものも None）。"""
    if value is None:
        return None
    if isinstance(value, datetime):
        return value.date()
    if isinstance(value, date):
        return value
    if not isinstance(value, str):
        return None
    s = re.sub(r"\s+", "", unicodedata.normalize("NFKC", value))
    if not s:
        return None
    m = _YMD.fullmatch(s)
    if m:
        return _make(int(m[1]), int(m[2]), int(m[3]))
    m = _ERA.fullmatch(s)
    if m:
        era = m[1].upper() if len(m[1]) == 1 else m[1]
        n = 1 if m[2] == "元" else int(m[2])
        return _make(_ERA_BASE[era] + n, int(m[3]), int(m[4]))
    return None


def _make(y: int, mo: int, d: int) -> date | None:
    try:
        return date(y, mo, d)
    except ValueError:
        return None
