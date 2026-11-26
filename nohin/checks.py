"""整合チェック。collect と report の両方から呼ばれる。"""

from __future__ import annotations

from collections import defaultdict
from pathlib import Path

from .classify import Location, NameInfo, expected_location, is_suspicious_other
from .config import Config, Delivery
from .model import (KIND_CL, KIND_DOC, KIND_OTHER, KIND_RV, LEVEL_ERROR, LEVEL_INFO, LEVEL_WARN,
                    STATUS_HEADERS, STATUS_KEYS, STATUSES, ST_DONE, ST_NA, ST_NG, TARGET_EXCLUDE, TARGETS,
                    Issue, Record)

MAX_PATH = 260
_LOC_LABEL = {"doc": "工程フォルダ直下", "review": "成果物/内部Rv・外部Rv"}


def check_all(records: list[Record], cfg: Config, delivery: Delivery, *, check_files: bool = True) -> list[Issue]:
    """records は enrich 済みであること。"""
    issues: list[Issue] = []
    active = [r for r in records if r.target != TARGET_EXCLUDE]
    issues += _check_each(active, cfg, check_files)
    for r in records:
        if r.target == TARGET_EXCLUDE and r.kind == KIND_OTHER and not is_suspicious_other(r):
            issues.append(Issue(LEVEL_INFO, r.project, "規則外（対象外）",
                                f"命名規則外のため対象外にしました: {r.rel}", (r.id,)))
    issues += _check_duplicates(active)
    issues += _check_pairs_and_rounds(active)
    issues += _check_missing(active, cfg, delivery)
    issues += _check_values_and_order(records)
    return issues


def _check_each(records: list[Record], cfg: Config, check_files: bool) -> list[Issue]:
    out = []
    for r in records:
        info: NameInfo = r.extra["name"]
        loc: Location = r.extra["location"]
        ids = (r.id,)
        if info.kind == KIND_OTHER:
            if is_suspicious_other(r):
                out.append(Issue(LEVEL_WARN, r.project, "規則外",
                                 f"納品物の置き場所にある Excel が命名規則に当てはまりません: {r.rel}", ids))
            else:
                out.append(Issue(LEVEL_INFO, r.project, "規則外",
                                 f"命名規則外のファイルです（手動で対象にした行）: {r.rel}", ids))
        else:
            exp = expected_location(info.kind)
            if loc.kind != exp:
                out.append(Issue(LEVEL_ERROR, r.project, "置き場所",
                                 f"{info.kind}が{_LOC_LABEL[exp]}以外にあります: {r.rel}", ids))
            folder_code = r.extra["folder_code"]
            if info.name_code and folder_code and info.name_code != folder_code:
                out.append(Issue(LEVEL_ERROR, r.project, "工程不一致",
                                 f"ファイル名の工程({info.name_code})とフォルダ({r.extra['folder']})が合いません: {r.filename}",
                                 ids))
            for prob in info.problems:
                out.append(Issue(LEVEL_ERROR, r.project, "名前の不整合", f"{prob}: {r.filename}", ids))
            if loc.kind == "review" and info.side and loc.side_dir and info.side != loc.side_dir:
                out.append(Issue(LEVEL_ERROR, r.project, "社内外不一致",
                                 f"「{loc.side_dir}」のフォルダに「{info.side}」のファイルがあります: {r.rel}", ids))
        if r.extra.get("fullwidth"):
            out.append(Issue(LEVEL_INFO, r.project, "全角括弧", f"全角の括弧を含みます: {r.filename}", ids))
        if check_files and r.src:
            if not Path(r.src).exists():
                out.append(Issue(LEVEL_ERROR, r.project, "元ファイルなし",
                                 f"元ファイルが見つかりません（移動・削除された可能性）: {r.src}", ids))
            if len(r.src) >= MAX_PATH:
                out.append(Issue(LEVEL_WARN, r.project, "パス長",
                                 f"パスが{MAX_PATH}文字以上です（{len(r.src)}文字）: {r.filename}", ids))
    return out


def _check_duplicates(records: list[Record]) -> list[Issue]:
    out = []
    docs = defaultdict(list)
    dests = defaultdict(list)
    for r in records:
        if r.kind == KIND_DOC:
            docs[(r.project, r.process, r.doc_name)].append(r)
        if r.dest_rel:
            dests[r.dest_rel].append(r)
    for (proj, code, name), rs in docs.items():
        if len(rs) > 1:
            out.append(Issue(LEVEL_ERROR, proj, "ドキュメント重複",
                             f"工程 {code} に{name}が{len(rs)}件あります（古い版は back/ へ）: "
                             + ", ".join(r.filename for r in rs), tuple(r.id for r in rs)))
    for dest, rs in dests.items():
        if len(rs) > 1:
            out.append(Issue(LEVEL_ERROR, rs[0].project, "納品先衝突",
                             f"納品先で同じ名前になります: {dest}", tuple(r.id for r in rs)))
    return out


def _check_pairs_and_rounds(records: list[Record]) -> list[Issue]:
    out = []
    pairs: dict[tuple, dict[str, list[Record]]] = defaultdict(lambda: {KIND_RV: [], KIND_CL: []})
    rounds: dict[tuple, list[Record]] = defaultdict(list)
    for r in records:
        if r.kind in (KIND_RV, KIND_CL) and r.side and r.round:
            pairs[(r.project, r.process, r.side, r.round)][r.kind].append(r)
            if r.kind == KIND_RV:
                rounds[(r.project, r.process, r.side)].append(r)
    for (proj, code, side, rnd), d in sorted(pairs.items()):
        label = f"工程 {code} {side} {rnd}回目"
        if d[KIND_RV] and not d[KIND_CL]:
            out.append(Issue(LEVEL_WARN, proj, "ペア不足",
                             f"{label}: レビュー記録表はあるがチェックリストがありません",
                             tuple(r.id for r in d[KIND_RV])))
        if d[KIND_CL] and not d[KIND_RV]:
            out.append(Issue(LEVEL_WARN, proj, "ペア不足",
                             f"{label}: チェックリストはあるがレビュー記録表がありません",
                             tuple(r.id for r in d[KIND_CL])))
    for (proj, code, side), rs in sorted(rounds.items()):
        have = {r.round for r in rs}
        missing = [n for n in range(1, max(have) + 1) if n not in have]
        if missing:
            out.append(Issue(LEVEL_WARN, proj, "回数の抜け",
                             f"工程 {code} {side}: レビュー記録表の {', '.join(f'{n}回目' for n in missing)} がありません",
                             tuple(r.id for r in rs)))
    return out


def _check_missing(records: list[Record], cfg: Config, delivery: Delivery) -> list[Issue]:
    out = []
    for proj in delivery.projects:
        mine = [r for r in records if r.project == proj.name]
        codes = list(proj.processes) or sorted({r.process for r in mine if r.process in cfg.processes})
        for code in codes:
            here = [r for r in mine if r.process == code]
            for rule in cfg.documents_for(code):
                if not any(r.kind == KIND_DOC and r.doc_name == rule.name for r in here):
                    out.append(Issue(LEVEL_WARN, proj.name, "ドキュメントなし",
                                     f"工程 {code} に{rule.name}がありません"))
            if not any(r.kind == KIND_RV for r in here):
                out.append(Issue(LEVEL_WARN, proj.name, "レビュー記録表なし",
                                 f"工程 {code} にレビュー記録表がありません"))
    return out


def _check_values_and_order(records: list[Record]) -> list[Issue]:
    out = []
    for r in records:
        ids = (r.id,)
        if r.target not in TARGETS:
            out.append(Issue(LEVEL_ERROR, r.project, "入力値",
                             f"No.{r.no} 対象の値「{r.target}」が不正です（{'/'.join(TARGETS)}）", ids))
        if r.target == TARGET_EXCLUDE:
            continue
        for k in STATUS_KEYS:
            v = getattr(r, k)
            if v not in STATUSES:
                out.append(Issue(LEVEL_ERROR, r.project, "入力値",
                                 f"No.{r.no} {STATUS_HEADERS[k]}の値「{v}」が不正です（{'/'.join(STATUSES)}）", ids))
            elif v == ST_NG:
                out.append(Issue(LEVEL_ERROR, r.project, "NG",
                                 f"No.{r.no} {STATUS_HEADERS[k]}が NG です: {r.filename}", ids))
        # 前の工程が済んでいないのに後の工程が済になっていないか
        for i, k in enumerate(STATUS_KEYS):
            if getattr(r, k) != ST_DONE:
                continue
            pending = [STATUS_HEADERS[p] for p in STATUS_KEYS[:i] if getattr(r, p) not in (ST_DONE, ST_NA)]
            if pending:
                out.append(Issue(LEVEL_WARN, r.project, "工程順序",
                                 f"No.{r.no} {STATUS_HEADERS[k]}が済ですが、{'・'.join(pending)}が済になっていません", ids))
                break
    return out


def sort_issues(issues: list[Issue]) -> list[Issue]:
    from .model import LEVEL_ORDER
    return sorted(issues, key=lambda i: (LEVEL_ORDER.get(i.level, 9), i.project, i.category, i.message))
