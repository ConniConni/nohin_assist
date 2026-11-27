"""プロジェクトルートを走査して、候補ファイルを集める（読み取りのみ）。"""

from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path

from .config import Config, Project
from .model import LEVEL_ERROR, LEVEL_INFO, LEVEL_WARN, Issue


@dataclass
class ScannedFile:
    project: Project
    path: Path
    rel: str  # プロジェクトルートからの相対パス（/ 区切り）


def _ignored(name: str, cfg: Config) -> bool:
    return any(p.search(name) for p in cfg.layout.ignore_patterns)


def scan_project(project: Project, cfg: Config) -> tuple[list[ScannedFile], list[Issue]]:
    files: list[ScannedFile] = []
    issues: list[Issue] = []
    root = project.root

    if not root.is_dir():
        issues.append(Issue(LEVEL_ERROR, project.name, "ルートなし",
                            f"プロジェクトルートが見つかりません: {root}"))
        return files, issues

    found_codes: dict[str, list[str]] = {}
    for entry in sorted(root.iterdir(), key=lambda p: p.name):
        if not entry.is_dir() or _ignored(entry.name, cfg):
            continue
        m = cfg.layout.process_dir_pattern.search(entry.name)
        if not m:
            continue
        code = m.group(1)
        found_codes.setdefault(code, []).append(entry.name)
        if code not in cfg.processes:
            issues.append(Issue(LEVEL_INFO, project.name, "未定義の工程",
                                f"工程の対応表にないフォルダのため読み飛ばしました: {entry.name}"))
            continue
        if project.processes and code not in project.processes:
            continue
        if entry.name != cfg.processes[code].folder:
            issues.append(Issue(LEVEL_INFO, project.name, "フォルダ名",
                                f"工程フォルダ名が対応表と異なります: {entry.name}（対応表: {cfg.processes[code].folder}）"))
        files.extend(_walk_process(project, entry, cfg))

    for code, names in found_codes.items():
        if len(names) > 1:
            issues.append(Issue(LEVEL_WARN, project.name, "工程フォルダ重複",
                                f"工程 {code} のフォルダが複数あります: {', '.join(names)}"))
    for code in project.processes:
        if code not in found_codes:
            issues.append(Issue(LEVEL_ERROR, project.name, "工程フォルダなし",
                                f"対象工程 {code}（{cfg.processes[code].folder}）のフォルダがありません"))
    return files, issues


def _walk_process(project: Project, proc_dir: Path, cfg: Config) -> list[ScannedFile]:
    """納品物の置き場所だけを見る: 工程フォルダ直下と、成果物/内部Rv・外部Rv（その下の日付フォルダを含む）。

    それ以外のフォルダ（作業用フォルダや back/ など）の下は管理表に載せない。
    """
    lay = cfg.layout
    out = _files_in(project, proc_dir, cfg)
    for deliv in _subdirs(proc_dir, cfg):
        if deliv.name != lay.deliverable_dir:
            continue
        for rv_dir in _subdirs(deliv, cfg):
            if rv_dir.name not in lay.review_dirs:
                continue
            out += _files_in(project, rv_dir, cfg)
            for date_dir in _subdirs(rv_dir, cfg):
                if lay.date_dir_pattern.search(date_dir.name):
                    out += _files_in(project, date_dir, cfg)
    return out


def _subdirs(d: Path, cfg: Config) -> list[Path]:
    return sorted((p for p in d.iterdir()
                   if p.is_dir() and p.name.lower() not in cfg.layout.excluded_dirs and not _ignored(p.name, cfg)),
                  key=lambda p: p.name)


def _files_in(project: Project, d: Path, cfg: Config) -> list[ScannedFile]:
    out = []
    for path in sorted((p for p in d.iterdir() if p.is_file()), key=lambda p: p.name):
        if _ignored(path.name, cfg):
            continue
        out.append(ScannedFile(project=project, path=path, rel=path.relative_to(project.root).as_posix()))
    return out
