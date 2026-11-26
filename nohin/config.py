"""config.toml（共通ルール）と delivery.toml（納品ごと）の読み込みと検証。"""

from __future__ import annotations

import re
import tomllib
from dataclasses import dataclass, field
from pathlib import Path


class ConfigError(Exception):
    """設定ファイルの内容に問題がある。"""


@dataclass(frozen=True)
class Process:
    code: str
    folder: str
    cl_name: str
    rv_name: str


@dataclass(frozen=True)
class DocRule:
    process: str
    name: str
    patterns: tuple[re.Pattern, ...]
    self_stamp: bool
    partner_stamp: bool


@dataclass(frozen=True)
class ReviewRule:
    kind: str  # "レビュー記録表" / "チェックリスト"
    patterns: tuple[re.Pattern, ...]
    self_stamp: bool
    partner_stamp_sides: tuple[str, ...]


@dataclass(frozen=True)
class Layout:
    process_dir_pattern: re.Pattern
    excluded_dirs: tuple[str, ...]
    deliverable_dir: str
    review_dirs: dict[str, str]
    date_dir_pattern: re.Pattern
    ignore_patterns: tuple[re.Pattern, ...]


@dataclass(frozen=True)
class ReviewTableConfig:
    """レビュー記録表の中身（指摘の表・承認欄）を読むための設定。文字は正規化して比較する。"""
    finding_labels: tuple[str, ...] = ("指摘事項", "指摘内容")
    action_labels: tuple[str, ...] = ("処置欄", "処置内容")
    date_labels: tuple[str, ...] = ("日付", "処置日")
    self_stamp_labels: tuple[str, ...] = ("捺印(自社)", "自社印")
    partner_stamp_labels: tuple[str, ...] = ("捺印(取引先)", "取引先印")
    no_labels: tuple[str, ...] = ("No.", "No", "NO.", "番号")
    approval_labels: tuple[str, ...] = ("承認",)
    not_required_marks: tuple[str, ...] = ("-", "ー", "―", "‐", "−", "—", "─")
    no_finding_marks: tuple[str, ...] = ("指摘なし", "以下余白")
    blank_stop_rows: int = 5
    header_search_rows: int = 200


@dataclass(frozen=True)
class Config:
    path: Path
    layout: Layout
    processes: dict[str, Process]
    documents: tuple[DocRule, ...]
    review_record: ReviewRule
    checklist: ReviewRule
    handler: str = ""      # 行ごとの自社印（対応者）の既定値
    approver: str = ""     # 社内レビュー記録表の承認者の既定値
    company_name: str = ""
    review_table: ReviewTableConfig = field(default_factory=ReviewTableConfig)

    def process_by_rv_name(self, name: str) -> Process | None:
        return next((p for p in self.processes.values() if p.rv_name == name), None)

    def documents_for(self, code: str) -> list[DocRule]:
        return [d for d in self.documents if d.process == code]


@dataclass(frozen=True)
class Project:
    name: str
    root: Path
    processes: tuple[str, ...] = ()  # 空なら「フォルダがある工程すべて」


@dataclass(frozen=True)
class Delivery:
    path: Path
    name: str
    work_dir: Path
    projects: tuple[Project, ...] = field(default_factory=tuple)

    @property
    def manifest_path(self) -> Path:
        return self.work_dir / f"納品管理表_{_safe_filename(self.name)}.xlsx"


def _safe_filename(name: str) -> str:
    return re.sub(r'[\\/:*?"<>|]', "_", name).strip() or "delivery"


def _load_toml(path: Path) -> dict:
    if not path.exists():
        raise ConfigError(f"設定ファイルが見つかりません: {path}")
    try:
        with path.open("rb") as f:
            return tomllib.load(f)
    except tomllib.TOMLDecodeError as e:
        raise ConfigError(f"{path.name} の書式が正しくありません: {e}") from e


def _compile(pattern: str, where: str, required_groups: tuple[str, ...] = ()) -> re.Pattern:
    try:
        rx = re.compile(pattern)
    except re.error as e:
        raise ConfigError(f"{where} の正規表現が正しくありません: {pattern!r} ({e})") from e
    missing = [g for g in required_groups if g not in rx.groupindex]
    if missing:
        raise ConfigError(f"{where} のパターンに名前付きグループ {missing} がありません: {pattern!r}")
    return rx


def _require(d: dict, key: str, where: str):
    if key not in d:
        raise ConfigError(f"{where} に `{key}` がありません")
    return d[key]


def _review_rule(d: dict, kind: str, where: str, groups: tuple[str, ...]) -> ReviewRule:
    pats = _require(d, "patterns", where)
    if not pats:
        raise ConfigError(f"{where}.patterns が空です")
    return ReviewRule(
        kind=kind,
        patterns=tuple(_compile(p, f"{where}.patterns", groups) for p in pats),
        self_stamp=bool(d.get("self_stamp", False)),
        partner_stamp_sides=tuple(d.get("partner_stamp_sides", [])),
    )


def load_config(path: str | Path) -> Config:
    path = Path(path).resolve()
    raw = _load_toml(path)

    lay = raw.get("layout", {})
    layout = Layout(
        process_dir_pattern=_compile(lay.get("process_dir_pattern", r"^(\d{3})\."), "layout.process_dir_pattern"),
        excluded_dirs=tuple(s.lower() for s in lay.get("excluded_dirs", ["back"])),
        deliverable_dir=lay.get("deliverable_dir", "成果物"),
        review_dirs=dict(lay.get("review_dirs", {"内部Rv": "社内", "外部Rv": "社外"})),
        date_dir_pattern=_compile(lay.get("date_dir_pattern", r"^\d{8}$"), "layout.date_dir_pattern"),
        ignore_patterns=tuple(_compile(p, "layout.ignore_patterns") for p in lay.get("ignore_patterns", [r"^~\$"])),
    )
    if layout.process_dir_pattern.groups < 1:
        raise ConfigError("layout.process_dir_pattern には工程番号を取り出す () が1つ必要です")

    processes: dict[str, Process] = {}
    for i, p in enumerate(_require(raw, "processes", "config")):
        where = f"processes[{i}]"
        proc = Process(
            code=str(_require(p, "code", where)),
            folder=_require(p, "folder", where),
            cl_name=_require(p, "cl_name", where),
            rv_name=_require(p, "rv_name", where),
        )
        if proc.code in processes:
            raise ConfigError(f"工程番号 {proc.code} が重複しています")
        processes[proc.code] = proc

    docs = []
    for i, d in enumerate(raw.get("documents", [])):
        where = f"documents[{i}]"
        code = str(_require(d, "process", where))
        if code not in processes:
            raise ConfigError(f"{where}.process={code} は processes にない工程番号です")
        docs.append(DocRule(
            process=code,
            name=_require(d, "name", where),
            patterns=tuple(_compile(p, f"{where}.patterns") for p in _require(d, "patterns", where)),
            self_stamp=bool(d.get("self_stamp", False)),
            partner_stamp=bool(d.get("partner_stamp", False)),
        ))

    stamp = raw.get("stamp", {})
    rt_raw = raw.get("review_table", {})
    defaults = ReviewTableConfig()
    rt_kwargs = {}
    for key in ("finding_labels", "action_labels", "date_labels", "self_stamp_labels", "partner_stamp_labels",
                "no_labels", "approval_labels", "not_required_marks", "no_finding_marks"):
        if key in rt_raw:
            v = rt_raw[key]
            if not isinstance(v, list) or not all(isinstance(x, str) for x in v):
                raise ConfigError(f"review_table.{key} は文字列のリストで書いてください")
            rt_kwargs[key] = tuple(v)
    for key in ("blank_stop_rows", "header_search_rows"):
        if key in rt_raw:
            v = rt_raw[key]
            if not isinstance(v, int) or v < 1:
                raise ConfigError(f"review_table.{key} は1以上の整数で書いてください")
            rt_kwargs[key] = v
    for key in ("finding_labels", "action_labels", "date_labels", "self_stamp_labels", "approval_labels"):
        if not rt_kwargs.get(key, getattr(defaults, key)):
            raise ConfigError(f"review_table.{key} が空です")

    return Config(
        path=path,
        layout=layout,
        processes=processes,
        documents=tuple(docs),
        review_record=_review_rule(_require(raw, "review_record", "config"), "レビュー記録表",
                                   "review_record", ("side", "round", "pname")),
        checklist=_review_rule(_require(raw, "checklist", "config"), "チェックリスト",
                               "checklist", ("code", "pname", "side", "round")),
        handler=stamp.get("handler", stamp.get("default_stamper", "")),
        approver=stamp.get("approver", ""),
        company_name=stamp.get("company_name", ""),
        review_table=ReviewTableConfig(**rt_kwargs),
    )


def load_delivery(path: str | Path, cfg: Config | None = None) -> Delivery:
    path = Path(path).resolve()
    raw = _load_toml(path)
    base = path.parent

    def resolve(p: str) -> Path:
        q = Path(p)
        return q if q.is_absolute() else (base / q)

    dv = _require(raw, "delivery", "delivery")
    name = _require(dv, "name", "delivery")
    work_dir = resolve(dv.get("work_dir", "work"))

    projects = []
    names = set()
    for i, p in enumerate(_require(raw, "projects", "delivery")):
        where = f"projects[{i}]"
        pname = _require(p, "name", where)
        if pname in names:
            raise ConfigError(f"案件名 {pname} が重複しています")
        names.add(pname)
        procs = tuple(str(c) for c in p.get("processes", []))
        if cfg is not None:
            unknown = [c for c in procs if c not in cfg.processes]
            if unknown:
                raise ConfigError(f"{where}.processes に未定義の工程番号があります: {unknown}")
        projects.append(Project(name=pname, root=resolve(_require(p, "root", where)), processes=procs))
    if not projects:
        raise ConfigError("delivery.toml に [[projects]] が1つもありません")
    return Delivery(path=path, name=name, work_dir=work_dir, projects=tuple(projects))


TOOL_DIR = Path(__file__).resolve().parent.parent


def find_config_for_inspect(explicit: str | None) -> Path:
    """inspect 用: 明示指定 → ツールのフォルダの config.toml → config.example.toml。"""
    if explicit:
        return Path(explicit)
    for c in (TOOL_DIR / "config.toml", TOOL_DIR / "config.example.toml"):
        if c.exists():
            return c
    raise ConfigError("config.toml が見つかりません")


def find_config(delivery_path: Path, explicit: str | None) -> Path:
    """config.toml を探す: 明示指定 → delivery.toml と同じ場所 → ツールのフォルダ。"""
    if explicit:
        return Path(explicit)
    candidates = [delivery_path.parent / "config.toml", Path(__file__).resolve().parent.parent / "config.toml"]
    for c in candidates:
        if c.exists():
            return c
    raise ConfigError("config.toml が見つかりません。config.example.toml をコピーして作成してください")
