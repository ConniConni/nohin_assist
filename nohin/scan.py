"""プロジェクトルートを走査して、候補ファイルを集める（読み取りのみ）。"""

from __future__ import annotations

import os
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
    out = []
    for dirpath, dirnames, filenames in os.walk(proc_dir):
        dirnames[:] = sorted(d for d in dirnames
                             if d.lower() not in cfg.layout.excluded_dirs and not _ignored(d, cfg))
        for fn in sorted(filenames):
            if _ignored(fn, cfg):
                continue
            path = Path(dirpath) / fn
            rel = path.relative_to(project.root).as_posix()
            out.append(ScannedFile(project=project, path=path, rel=rel))
    return out
