"""ファイル名と置き場所から、種別・工程・社内外・回数を判定する。"""

from __future__ import annotations

import unicodedata
from dataclasses import dataclass, field

from .config import Config, DocRule
from .model import (KIND_CL, KIND_DOC, KIND_OTHER, KIND_RV, ST_NA, ST_TODO, TARGET_AUTO, TARGET_CHECK,
                    TARGET_EXCLUDE, Record)

FULLWIDTH_BRACKETS = ("（", "）")


def normalize(name: str) -> str:
    """全角の括弧・数字などを半角にそろえる（判定用）。"""
    return unicodedata.normalize("NFKC", name)


@dataclass
class NameInfo:
    kind: str = KIND_OTHER
    doc_rule: DocRule | None = None
    side: str = ""
    round: int | None = None
    name_code: str | None = None   # ファイル名から読み取れる工程番号
    pname: str = ""
    problems: list[str] = field(default_factory=list)


def classify_name(filename: str, cfg: Config) -> NameInfo:
    n = normalize(filename)

    for rule in cfg.documents:
        if any(p.search(n) for p in rule.patterns):
            return NameInfo(kind=KIND_DOC, doc_rule=rule, name_code=rule.process)

    for p in cfg.review_record.patterns:
        m = p.search(n)
        if m:
            info = NameInfo(kind=KIND_RV, side=m["side"], round=int(m["round"]), pname=m["pname"])
            proc = cfg.process_by_rv_name(m["pname"])
            if proc:
                info.name_code = proc.code
            else:
                info.problems.append(f"工程名「{m['pname']}」が工程の対応表(rv_name)にありません")
            return info

    for p in cfg.checklist.patterns:
        m = p.search(n)
        if m:
            info = NameInfo(kind=KIND_CL, side=m["side"], round=int(m["round"]), pname=m["pname"],
                            name_code=m["code"])
            proc = cfg.processes.get(m["code"])
            if proc is None:
                info.problems.append(f"工程番号 {m['code']} が工程の対応表にありません")
            elif proc.cl_name != m["pname"]:
                info.problems.append(
                    f"工程名「{m['pname']}」が工程番号 {m['code']} と合いません（正: {proc.cl_name}）")
            return info

    return NameInfo()


@dataclass
class Location:
    kind: str            # "doc"（工程直下）/ "review"（成果物/内部Rv・外部Rv）/ "other"
    side_dir: str = ""   # フォルダから読み取れる社内外
    date_dir: str = ""


def analyze_location(parts: tuple[str, ...], cfg: Config) -> Location:
    """parts は工程フォルダからの相対パス（ファイル名を含む）。"""
    lay = cfg.layout
    if len(parts) == 1:
        return Location("doc")
    if len(parts) in (3, 4) and parts[0] == lay.deliverable_dir and parts[1] in lay.review_dirs:
        side = lay.review_dirs[parts[1]]
        if len(parts) == 3:
            return Location("review", side_dir=side)
        if lay.date_dir_pattern.search(parts[2]):
            return Location("review", side_dir=side, date_dir=parts[2])
    return Location("other")


def enrich(rec: Record, cfg: Config) -> Record:
    """相対パスとファイル名から、チェック用の付加情報（extra）を計算する。"""
    parts = tuple(p for p in rec.rel.split("/") if p)
    folder = parts[0] if parts else ""
    m = cfg.layout.process_dir_pattern.search(folder)
    loc = analyze_location(parts[1:], cfg) if len(parts) >= 2 else Location("other")
    info = classify_name(rec.filename, cfg)
    rec.extra = {
        "folder": folder,
        "folder_code": m.group(1) if m else "",
        "location": loc,
        "name": info,
        "fullwidth": any(b in rec.filename for b in FULLWIDTH_BRACKETS),
    }
    return rec


EXCEL_EXTS = (".xlsx", ".xlsm", ".xls")


def is_suspicious_other(rec: Record) -> bool:
    """規則外のうち「名前を間違えた納品物かもしれない」もの: 納品物の置き場所にある Excel ファイル。"""
    loc: Location = rec.extra["location"]
    return loc.kind in ("doc", "review") and rec.filename.lower().endswith(EXCEL_EXTS)


def in_scope(rec: Record) -> bool:
    """管理表に載せるか。ドキュメントは工程フォルダ直下にあるものだけを載せる。

    レビューフォルダには、レビューに使ったドキュメントの写しが同じ名前で置かれることがあるため、
    そこにあるドキュメント名のファイルは載せない（重複にも数えない）。
    """
    loc: Location = rec.extra["location"]
    if loc.kind == "doc":
        return True
    return loc.kind == "review" and rec.kind != KIND_DOC


def expected_location(kind: str) -> str | None:
    return {KIND_DOC: "doc", KIND_RV: "review", KIND_CL: "review"}.get(kind)


def build_record(project_name: str, rel: str, src: str, cfg: Config, now: str) -> Record:
    """走査したファイルから、管理表の新しい行を作る。"""
    filename = rel.rsplit("/", 1)[-1]
    rec = Record(project=project_name, filename=filename, rel=rel, src=src,
                 id=f"{project_name}/{rel}", added_at=now)
    enrich(rec, cfg)
    info: NameInfo = rec.extra["name"]
    loc: Location = rec.extra["location"]
    folder = rec.extra["folder"]
    rec.process = rec.extra["folder_code"]
    rec.kind = info.kind

    if info.kind == KIND_OTHER:
        # 基本は対象外。納品物の置き場所にある Excel だけは名前の誤りの可能性があるので人に確認してもらう
        rec.target = TARGET_CHECK if is_suspicious_other(rec) else TARGET_EXCLUDE
        rec.st_copy = ST_TODO
        return rec

    # 納品先の工程フォルダ名は対応表の名前にそろえる（作業元のフォルダ名が多少違っても同じ場所へ）
    proc = cfg.processes.get(rec.process)
    dest_folder = proc.folder if proc else folder
    if info.kind == KIND_DOC:
        rule = info.doc_rule
        rec.doc_name = rule.name
        self_stamp, partner = rule.self_stamp, rule.partner_stamp
        rec.dest_rel = f"{project_name}/{dest_folder}/{filename}"
    else:
        rr = cfg.review_record if info.kind == KIND_RV else cfg.checklist
        rec.doc_name = rr.kind
        rec.side, rec.round = info.side, info.round
        self_stamp, partner = rr.self_stamp, info.side in rr.partner_stamp_sides
        rec.dest_rel = f"{project_name}/{dest_folder}/{cfg.layout.deliverable_dir}/{filename}"

    rec.st_self = ST_TODO if self_stamp else ST_NA
    rec.st_send = rec.st_partner = ST_TODO if partner else ST_NA
    rec.st_copy = ST_TODO
    # 指摘展開は社外レビュー記録表だけ（指摘を書いた時点で取引先に展開する）
    rec.st_share = ST_TODO if info.kind == KIND_RV and partner else ST_NA
    rec.teams_rel = rec.dest_rel if partner else ""
    rec.handler = cfg.handler if self_stamp else ""
    # 承認欄に自社が押すのは、取引先に展開しない（社内）レビュー記録表だけ
    rec.approver = cfg.approver if info.kind == KIND_RV and not partner else ""
    side_ok = not (loc.kind == "review" and loc.side_dir and info.side and loc.side_dir != info.side)
    code_ok = not (info.name_code and rec.process and info.name_code != rec.process)
    good = loc.kind == expected_location(info.kind) and not info.problems and side_ok and code_ok
    rec.target = TARGET_AUTO if good else TARGET_CHECK
    return rec
