"""納品コピーで使うファイル操作: ハッシュ、一時ファイル経由のコピー、納品先パスの解決、名前の正規化。"""

from __future__ import annotations

import hashlib
import os
import re
import shutil
import unicodedata
from dataclasses import dataclass
from pathlib import Path

from .classify import Location
from .config import Config, Delivery
from .model import Record

CHUNK = 1024 * 1024
TMP_SUFFIX = ".nohin-tmp"
_DRIVE = re.compile(r"^[A-Za-z]:")


@dataclass
class HashStats:
    files: int = 0
    bytes: int = 0


STATS = HashStats()
_cache: dict[tuple[str, int, int], str] = {}


def reset_stats() -> None:
    STATS.files = STATS.bytes = 0
    _cache.clear()


def sha256_of(path: Path) -> str:
    """ファイルの SHA-256（16進）。1MB ずつ読む。同じ実行の中で変わっていないファイルは読み直さない。"""
    st = path.stat()
    key = (str(path), st.st_size, st.st_mtime_ns)
    if key in _cache:
        return _cache[key]
    h = hashlib.sha256()
    with path.open("rb") as f:
        while chunk := f.read(CHUNK):
            h.update(chunk)
    STATS.files += 1
    STATS.bytes += st.st_size
    _cache[key] = h.hexdigest()
    return _cache[key]


def norm_key(rel: str) -> str:
    """名前の比較用: 区切りを / に、文字の表し方を NFC にそろえ、大文字・小文字を区別しない。"""
    s = unicodedata.normalize("NFC", rel.replace("\\", "/"))
    return "/".join(p for p in s.split("/") if p).casefold()


def dest_parts(rec: Record) -> list[str] | None:
    """納品先(相対)を部品に分ける。空なら []、納品先の外を指すなら None。"""
    rel = rec.dest_rel.strip().replace("\\", "/")
    if not rel:
        return []
    if rel.startswith("/") or _DRIVE.match(rel):
        return None
    parts = [p for p in rel.split("/") if p]
    if any(p in (".", "..") for p in parts) or len(parts) < 2:
        return None
    if norm_key(parts[0]) != norm_key(rec.project):  # 先頭はその行の案件名
        return None
    return parts


def dest_path(delivery: Delivery, rec: Record) -> Path | None:
    """納品先の絶対パス。dest_root がない・納品先(相対)が空・不正なら None。"""
    if delivery.dest_root is None:
        return None
    parts = dest_parts(rec)
    if not parts:
        return None
    return delivery.dest_root.joinpath(*parts)


def suggest_dest_rel(rec: Record, cfg: Config) -> str:
    """置き場所から納品先(相対)の候補を作る（表示用。管理表には書かない）。決まらなければ空。"""
    loc: Location | None = rec.extra.get("location")
    folder = rec.extra.get("folder", "")
    if loc is None or not folder:
        return ""
    proc = cfg.processes.get(rec.extra.get("folder_code", ""))
    dest_folder = proc.folder if proc else folder
    if loc.kind == "doc":
        return f"{rec.project}/{dest_folder}/{rec.filename}"
    if loc.kind == "review":
        return f"{rec.project}/{dest_folder}/{cfg.layout.deliverable_dir}/{rec.filename}"
    return ""


def is_editing(path: Path) -> bool:
    """Excel で開いているときにできる ~$ ファイルがあるか。"""
    return any((path.parent / n).exists() for n in (f"~${path.name}", f"~${path.name[2:]}"))


class CopyError(Exception):
    """コピーまたは照合に失敗した（納品先は元のまま）。"""


def copy_verified(src: Path, dest: Path) -> str:
    """src を dest へコピーしてハッシュを照合し、一致したハッシュを返す。

    同じフォルダの一時ファイルに書いてから置き換えるので、途中で失敗しても dest は元のまま。
    """
    dest.parent.mkdir(parents=True, exist_ok=True)
    tmp = dest.with_name(dest.name + TMP_SUFFIX)
    try:
        shutil.copyfile(src, tmp)
        try:
            shutil.copystat(src, tmp)  # 共有フォルダでは属性のコピーに失敗することがある（中身は問題ない）
        except OSError:
            pass
        expected, got = sha256_of(src), sha256_of(tmp)
        if expected != got:
            raise CopyError("コピー後のハッシュが元ファイルと一致しません")
        os.replace(tmp, dest)
        return got
    except BaseException:
        try:
            tmp.unlink()
        except OSError:
            pass
        raise


def backup_file(path: Path, backup_root: Path, parts: list[str]) -> Path:
    """path を backup_root の下に、納品先と同じ相対パスで退避する。"""
    out = backup_root.joinpath(*parts)
    out.parent.mkdir(parents=True, exist_ok=True)
    shutil.copy2(path, out)
    if sha256_of(path) != sha256_of(out):
        raise CopyError("バックアップのハッシュが一致しません")
    return out
