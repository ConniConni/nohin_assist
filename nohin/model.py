"""管理表の1行（Record）と列定義、ステータス値。"""

from __future__ import annotations

from dataclasses import dataclass, field, fields

# 種別
KIND_DOC = "ドキュメント"
KIND_RV = "レビュー記録表"
KIND_CL = "チェックリスト"
KIND_OTHER = "規則外"

# 対象
TARGET_AUTO = "自動"
TARGET_MANUAL = "手動"
TARGET_EXCLUDE = "除外"
TARGET_CHECK = "要確認"
TARGETS = (TARGET_AUTO, TARGET_MANUAL, TARGET_EXCLUDE, TARGET_CHECK)

# ステータス
ST_TODO = "未"
ST_DONE = "済"
ST_NG = "NG"
ST_NA = "不要"
STATUSES = (ST_TODO, ST_DONE, ST_NG, ST_NA)

# 工程ステータス列（この順に進む）
STATUS_KEYS = ("st_self", "st_send", "st_partner", "st_copy")


@dataclass
class Record:
    # --- 表示列 ---
    no: int | None = None
    project: str = ""
    process: str = ""          # 置かれている工程フォルダの工程番号
    kind: str = ""
    doc_name: str = ""
    side: str = ""             # 社内 / 社外
    round: int | None = None   # 回数
    filename: str = ""
    target: str = TARGET_AUTO
    st_self: str = ST_NA
    st_send: str = ST_NA
    st_partner: str = ST_NA
    st_copy: str = ST_TODO
    warning: str = ""
    note: str = ""
    # --- 詳細列（グループ化して折りたたむ） ---
    id: str = ""               # 案件名/相対パス（行の識別子）
    src: str = ""              # 元ファイルの絶対パス
    rel: str = ""              # プロジェクトルートからの相対パス（/ 区切り）
    dest_rel: str = ""         # 納品先（相対）
    teams_rel: str = ""        # Teams 格納先（相対）
    stamper: str = ""
    stamp_sheet: str = ""
    stamp_cell: str = ""
    date_cell: str = ""
    stamp_date: str = ""
    self_at: str = ""
    send_at: str = ""
    send_size: str = ""
    send_mtime: str = ""
    pull_at: str = ""
    copy_at: str = ""
    copy_hash: str = ""
    backup: str = ""
    added_at: str = ""
    # --- 管理表には書かない作業用の値 ---
    extra: dict = field(default_factory=dict, compare=False, repr=False)


# (属性名, 見出し, 列幅, 詳細列か)
COLUMNS: list[tuple[str, str, int, bool]] = [
    ("no", "No", 5, False),
    ("project", "案件", 14, False),
    ("process", "工程", 6, False),
    ("kind", "種別", 13, False),
    ("doc_name", "文書名", 14, False),
    ("side", "社内外", 7, False),
    ("round", "回数", 6, False),
    ("filename", "ファイル名", 48, False),
    ("target", "対象", 8, False),
    ("st_self", "自社捺印", 9, False),
    ("st_send", "取引先送付", 10, False),
    ("st_partner", "取引先捺印", 10, False),
    ("st_copy", "納品コピー", 10, False),
    ("warning", "警告", 24, False),
    ("note", "備考", 24, False),
    ("id", "ID", 40, True),
    ("src", "元パス", 60, True),
    ("rel", "相対パス", 48, True),
    ("dest_rel", "納品先(相対)", 48, True),
    ("teams_rel", "Teams格納先(相対)", 48, True),
    ("stamper", "押印者", 10, True),
    ("stamp_sheet", "押印シート", 10, True),
    ("stamp_cell", "押印セル", 10, True),
    ("date_cell", "日付セル", 10, True),
    ("stamp_date", "印影日付", 11, True),
    ("self_at", "自社捺印日時", 16, True),
    ("send_at", "送付日時", 16, True),
    ("send_size", "送付時サイズ", 12, True),
    ("send_mtime", "送付時更新日時", 16, True),
    ("pull_at", "引き戻し日時", 16, True),
    ("copy_at", "納品コピー日時", 16, True),
    ("copy_hash", "ハッシュ", 20, True),
    ("backup", "バックアップ先", 40, True),
    ("added_at", "追加日時", 16, True),
]

HEADER_TO_ATTR = {h: a for a, h, _, _ in COLUMNS}
ATTR_TO_HEADER = {a: h for a, h, _, _ in COLUMNS}
STATUS_HEADERS = {k: ATTR_TO_HEADER[k] for k in STATUS_KEYS}

_RECORD_FIELDS = {f.name for f in fields(Record)}
assert all(a in _RECORD_FIELDS for a, *_ in COLUMNS)


@dataclass
class Issue:
    level: str        # エラー / 警告 / 情報
    project: str
    category: str     # 短い区分名
    message: str
    row_ids: tuple[str, ...] = ()
    confirmed: str = ""  # 人が「確認済」にした印

    @property
    def key(self) -> str:
        """確認済マークを引き継ぐための識別子。"""
        return f"{self.category}|{self.project}|{self.message}"


LEVEL_ERROR = "エラー"
LEVEL_WARN = "警告"
LEVEL_INFO = "情報"
LEVEL_ORDER = {LEVEL_ERROR: 0, LEVEL_WARN: 1, LEVEL_INFO: 2}
