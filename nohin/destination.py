"""納品先のチェック。dest_root があれば collect・report・deliver のたびに同じ警告を出す。

警告の文には日時・ハッシュ・件数など変わる値を入れない（「確認済」を引き継ぐため）。
"""

from __future__ import annotations

import os
from pathlib import Path

from .checks import MAX_PATH
from .config import Config, Delivery
from .fileops import TMP_SUFFIX, dest_parts, dest_path, norm_key, sha256_of
from .model import (LEVEL_INFO, LEVEL_WARN, ST_DONE, TARGET_AUTO, TARGET_CHECK, TARGET_EXCLUDE, TARGET_MANUAL,
                    Issue, Record)

WHOLE = "全体"


def _ignored(name: str, cfg: Config) -> bool:
    return any(p.search(name) for p in cfg.layout.ignore_patterns)


def _same(a: Path, b: Path, verify: bool) -> bool:
    if verify:
        return sha256_of(a) == sha256_of(b)
    return a.stat().st_size == b.stat().st_size


def check_destination(records: list[Record], cfg: Config, delivery: Delivery) -> list[Issue]:
    root = delivery.dest_root
    if root is None:
        return []
    if not root.is_dir():
        return [Issue(LEVEL_INFO, WHOLE, "納品先のフォルダなし",
                      f"納品先のフォルダがまだありません（deliver の check・overwrite で作られます）: {root}")]
    verify = cfg.verify_hash_on_report
    issues: list[Issue] = []
    known: set[str] = set()
    for r in records:
        if r.target == TARGET_EXCLUDE:
            continue
        parts = dest_parts(r)
        if parts is None:
            issues.append(Issue(LEVEL_WARN, r.project, "納品先が不正",
                                f"納品先(相対)が案件「{r.project}」のフォルダの外を指しています: {r.dest_rel}", (r.id,)))
            continue
        if parts:
            known.add(norm_key("/".join(parts)))
        issues += _check_row(r, delivery, verify)
    issues += _check_strays(records, cfg, delivery, known)
    return issues


def _check_row(r: Record, delivery: Delivery, verify: bool) -> list[Issue]:
    dest = dest_path(delivery, r)
    if dest is None or r.target == TARGET_CHECK:
        return []
    ids = (r.id,)
    out: list[Issue] = []
    if len(str(dest)) >= MAX_PATH:
        out.append(Issue(LEVEL_WARN, r.project, "パス長",
                         f"納品先のパスが{MAX_PATH}文字以上です（{len(str(dest))}文字）: {r.dest_rel}", ids))
    if r.st_copy != ST_DONE:
        return out
    if not dest.is_file():
        return out + [Issue(LEVEL_WARN, r.project, "納品先なし",
                            f"納品コピーが済ですが、納品先にファイルがありません: {r.dest_rel}", ids)]
    src = Path(r.src) if r.src else None
    src_ok = src is not None and src.is_file()
    try:
        if r.target == TARGET_AUTO and r.copy_hash:
            if verify:
                if sha256_of(dest) != r.copy_hash:
                    out.append(Issue(LEVEL_WARN, r.project, "納品先の内容違い",
                                     f"納品先のファイルが、納品コピーしたときと違います: {r.dest_rel}", ids))
                elif src_ok and sha256_of(src) != r.copy_hash:
                    out.append(Issue(LEVEL_WARN, r.project, "納品後に原本が更新",
                                     f"納品コピーのあとに原本が変わっています（overwrite で置き直してください）: {r.filename}",
                                     ids))
            elif src_ok and not _same(src, dest, False):
                out.append(Issue(LEVEL_WARN, r.project, "納品先の内容違い",
                                 f"納品先と原本のサイズが違います（ハッシュの照合は設定で止めています）: {r.dest_rel}", ids))
        elif r.target in (TARGET_AUTO, TARGET_MANUAL) and src_ok and not _same(src, dest, verify):
            # 手で置いた行（ハッシュの記録がない行を含む）は、原本と比べる
            out.append(Issue(LEVEL_WARN, r.project, "手動配置の内容違い",
                             f"納品先のファイルが元ファイルと違います: {r.filename}", ids))
    except OSError as e:
        out.append(Issue(LEVEL_WARN, r.project, "納品先の読取不可",
                         f"納品先か原本を読めません（{type(e).__name__}）: {r.dest_rel}", ids))
    return out


def _check_strays(records: list[Record], cfg: Config, delivery: Delivery, known: set[str]) -> list[Issue]:
    out: list[Issue] = []
    # 納品先(相対)が空の行（同じ名前のファイルが納品先にあれば案内する）
    blank: dict[tuple[str, str], list[Record]] = {}
    for r in records:
        if r.target in (TARGET_AUTO, TARGET_MANUAL) and not r.dest_rel.strip():
            blank.setdefault((norm_key(r.project), norm_key(r.filename)), []).append(r)
    for proj in delivery.projects:
        base = delivery.dest_root / proj.name
        if not base.is_dir():
            continue
        for dirpath, dirnames, filenames in os.walk(base):
            dirnames[:] = sorted(d for d in dirnames if not _ignored(d, cfg))
            for fn in sorted(filenames):
                path = Path(dirpath) / fn
                rel = f"{proj.name}/{path.relative_to(base).as_posix()}"
                if fn.endswith(TMP_SUFFIX):
                    out.append(Issue(LEVEL_WARN, proj.name, "一時ファイルの残り",
                                     f"コピーの途中で止まったときの一時ファイルが残っています（消してください）: {rel}"))
                    continue
                if _ignored(fn, cfg) or norm_key(rel) in known:
                    continue
                rows = blank.get((norm_key(proj.name), norm_key(fn)), [])
                hint = "（同じ名前の行の納品先(相対)が空です）" if rows else ""
                out.append(Issue(LEVEL_WARN, proj.name, "管理表にないファイル",
                                 f"納品先にあるのに管理表にないファイルです: {rel}{hint}", tuple(r.id for r in rows)))
    return out
