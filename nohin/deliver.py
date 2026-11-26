"""④ 納品コピー（deliver）と、納品コピーのバックアップからの復元（restore）。

- list: 一覧を出すだけ。作業元・納品先・管理表のどれも変更しない（ログも書かない）。
- check: 内容の違う同名ファイルが1件でもあれば、1件もコピーせずに止める。
- overwrite: 内容の違う既存ファイルを _backup に退避してから上書きする。
作業元の原本は読むだけで、restore も含めて一切書き込まない。
"""

from __future__ import annotations

import shutil
import time
from dataclasses import dataclass, field
from datetime import datetime
from pathlib import Path

from .config import ConfigError, Delivery
from .fileops import (STATS, CopyError, backup_file, copy_verified, dest_parts, dest_path, is_editing, norm_key,
                      reset_stats, sha256_of, suggest_dest_rel)
from .manifest import ensure_not_open, read_manifest, save_manifest
from .model import (LEVEL_ERROR, STATUS_HEADERS, STATUS_KEYS, ST_DONE, ST_NA, ST_NG, ST_TODO, TARGET_AUTO,
                    TARGET_CHECK, TARGET_EXCLUDE, TARGET_MANUAL, Issue, Record)
from .refresh import refresh

MODES = ("list", "check", "overwrite")
MAX_LIST = 30
BACKUP_LABEL = "納品コピー"
RESTORE_LABEL = "restore"
BLOCKING_CATEGORIES = ("納品先衝突", "元ファイルなし")
PREREQ_KEYS = STATUS_KEYS[:-1]  # 納品コピーより前の工程

# 分類
CAT_COPY = "コピー対象"
CAT_MANUAL = "手で配置"
CAT_NO_DEST = "納品先が空"
CAT_BAD_DEST = "納品先が不正"
CAT_BLOCKED = "コピーできない"
CAT_PENDING = "前の工程が未"
CAT_CHECK = "要確認"
CAT_NOT_NEEDED = "納品コピーが不要"
CAT_EXCLUDE = "除外"

# 納品先の状態
NEW, SAME, DIFF = "新規", "同じ内容", "内容が違う"


@dataclass
class Item:
    rec: Record
    cat: str
    state: str = ""           # NEW / SAME / DIFF（コピー対象・手で配置）
    dest: Path | None = None
    note: str = ""            # 理由・候補など
    hash: str = ""            # SAME のときの納品先のハッシュ


@dataclass
class Output:
    """画面に出し、同じ内容をログにも残す。"""
    lines: list[str] = field(default_factory=list)

    def __call__(self, text: str = "") -> None:
        print(text)
        self.lines.append(text)

    def save(self, work_dir: Path, label: str, ts: str) -> Path:
        path = work_dir / "_log" / f"{label}_{ts}.txt"
        path.parent.mkdir(parents=True, exist_ok=True)
        path.write_text("\n".join(self.lines) + "\n", encoding="utf-8")
        return path


def _now() -> str:
    return datetime.now().strftime("%Y-%m-%d %H:%M")


def _stamp() -> str:
    return datetime.now().strftime("%Y%m%d-%H%M%S")


def _require_dest(delivery: Delivery) -> Path:
    if delivery.dest_root is None:
        raise ConfigError(f"delivery.toml の [delivery] に dest_root（納品先のフォルダ）を書いてください: {delivery.path}")
    return delivery.dest_root


def _no(rec: Record) -> str:
    return f"No.{rec.no}" if rec.no is not None else "No.-"


# ---------------------------------------------------------------- 分類

def _dest_state(src: Path, dest: Path) -> tuple[str, str]:
    if not dest.is_file():
        return NEW, ""
    h = sha256_of(dest)
    return (SAME if h == sha256_of(src) else DIFF), h


def classify_rows(records: list[Record], cfg, delivery: Delivery, issues: list[Issue]) -> list[Item]:
    """管理表の各行を、コピーするか・なぜ飛ばすかに分ける（records は enrich 済み）。"""
    blocked: dict[str, list[str]] = {}
    for i in issues:
        if i.level == LEVEL_ERROR and i.category in BLOCKING_CATEGORIES:
            for rid in i.row_ids:
                blocked.setdefault(rid, []).append(i.category)
    items: list[Item] = []
    for r in records:
        if r.target == TARGET_EXCLUDE:
            items.append(Item(r, CAT_EXCLUDE))
            continue
        if r.target == TARGET_CHECK:
            items.append(Item(r, CAT_CHECK))
            continue
        if r.target not in (TARGET_AUTO, TARGET_MANUAL):
            items.append(Item(r, CAT_BLOCKED, note=f"対象の値「{r.target}」が不正"))
            continue
        if r.st_copy == ST_NA:
            items.append(Item(r, CAT_NOT_NEEDED))
            continue
        parts = dest_parts(r)
        src = Path(r.src) if r.src else None
        if parts is None:
            items.append(Item(r, CAT_BAD_DEST, note=r.dest_rel))
            continue
        if not parts:
            cat = CAT_MANUAL if r.target == TARGET_MANUAL else CAT_NO_DEST
            sug = suggest_dest_rel(r, cfg)
            items.append(Item(r, cat, note=f"候補: {sug}" if sug else "候補なし"))
            continue
        dest = delivery.dest_root.joinpath(*parts)
        if r.target == TARGET_MANUAL:
            it = Item(r, CAT_MANUAL, dest=dest)
            if not dest.is_file():
                it.state = "なし"
            elif src is not None and src.is_file():
                try:
                    it.state = _dest_state(src, dest)[0]
                except OSError as e:
                    it.state, it.note = "読めない", type(e).__name__
            else:
                it.state = "あり（元ファイルなし）"
            items.append(it)
            continue
        pending = [STATUS_HEADERS[k] for k in PREREQ_KEYS if getattr(r, k) not in (ST_DONE, ST_NA)]
        if pending:
            items.append(Item(r, CAT_PENDING, note="・".join(pending)))
            continue
        reasons = list(blocked.get(r.id, []))
        if src is None or not src.is_file():
            if "元ファイルなし" not in reasons:
                reasons.append("元ファイルなし")
        elif is_editing(src):
            reasons.append("元ファイルが Excel で開かれている")
        if reasons:
            items.append(Item(r, CAT_BLOCKED, dest=dest, note="・".join(reasons)))
            continue
        try:
            state, h = _dest_state(src, dest)
        except OSError as e:
            items.append(Item(r, CAT_BLOCKED, dest=dest, note=f"読めない（{type(e).__name__}: {e}）"))
            continue
        items.append(Item(r, CAT_COPY, state=state, dest=dest, hash=h))
    return items


# ---------------------------------------------------------------- 表示

def print_plan(items: list[Item], out: Output) -> None:
    by: dict[str, list[Item]] = {}
    for it in items:
        by.setdefault(it.cat, []).append(it)
    copy = by.get(CAT_COPY, [])
    counts = {s: sum(1 for i in copy if i.state == s) for s in (NEW, SAME, DIFF)}
    out(f"コピー対象（自動）: {len(copy)} 件")
    out(f"  {NEW} {counts[NEW]} 件 / {SAME} {counts[SAME]} 件 / {DIFF} {counts[DIFF]} 件")
    for it in [i for i in copy if i.state == DIFF]:
        out(f"  {_no(it.rec)}  {it.rec.dest_rel}  → {DIFF}（納品先に違う内容の同名ファイルがあります）")
    news = [i for i in copy if i.state == NEW]
    for it in news[:MAX_LIST]:
        out(f"  {_no(it.rec)}  {it.rec.dest_rel}  → {NEW}")
    if len(news) > MAX_LIST:
        out(f"  …ほか {len(news) - MAX_LIST} 件")

    manual = by.get(CAT_MANUAL, [])
    if manual:
        out()
        out(f"手で配置（手動）: {len(manual)} 件")
        for it in manual:
            if it.dest is None:
                out(f"  {_no(it.rec)}  {it.rec.filename}  納品先(相対)が空です")
                out(f"         {it.note}")
            else:
                out(f"  {_no(it.rec)}  {it.rec.filename}  納品先: {it.state}" + (f"（{it.note}）" if it.note else ""))

    for cat, title in ((CAT_NO_DEST, "納品先(相対)が空（自動）"), (CAT_BAD_DEST, "納品先(相対)が不正（案件のフォルダの外を指している）"),
                       (CAT_BLOCKED, "コピーできない")):
        rows = by.get(cat, [])
        if not rows:
            continue
        out()
        out(f"{title}: {len(rows)} 件")
        for it in rows:
            out(f"  {_no(it.rec)}  {it.rec.filename}  {it.note}")

    out()
    out("飛ばした行:")
    pending = by.get(CAT_PENDING, [])
    shown = f"（最初の{MAX_LIST}件を表示）" if len(pending) > MAX_LIST else ""
    out(f"  {CAT_PENDING} {len(pending)} 件{shown}")
    for it in pending[:MAX_LIST]:
        out(f"    {_no(it.rec)}  {it.rec.filename}  （未: {it.note}）")
    check = by.get(CAT_CHECK, [])
    out(f"  {CAT_CHECK} {len(check)} 件" + ("  " + "、".join(f"{_no(i.rec)} {i.rec.filename}" for i in check) if check else ""))
    out(f"  {CAT_EXCLUDE} {len(by.get(CAT_EXCLUDE, []))} 件")
    if by.get(CAT_NOT_NEEDED):
        out(f"  {CAT_NOT_NEEDED} {len(by[CAT_NOT_NEEDED])} 件")


def _print_stats(out: Output, started: float) -> None:
    out(f"処理時間: {time.monotonic() - started:.1f} 秒（ハッシュを計算したファイル: {STATS.files} 件・"
        f"合計 {STATS.bytes / 1024 / 1024:.1f} MB）")


# ---------------------------------------------------------------- deliver

def _mark_done(rec: Record, h: str, now: str) -> None:
    rec.st_copy, rec.copy_at, rec.copy_hash = ST_DONE, now, h


def run_deliver(cfg, delivery: Delivery, mode: str) -> int:
    started = time.monotonic()
    reset_stats()
    _require_dest(delivery)
    out = Output()
    path = delivery.manifest_path
    records, confirmed = read_manifest(path)
    now = _now()
    _, issues = refresh(records, cfg, delivery, [], confirmed, now)
    items = classify_rows(records, cfg, delivery, issues)

    out(f"納品コピー（モード: {mode}）")
    out(f"管理表: {path}")
    out(f"納品先: {delivery.dest_root}")
    out()
    print_plan(items, out)
    out()

    if mode == "list":
        out("※ list なので何も変更していません。コピーするときは check か overwrite で実行してください。")
        _print_stats(out, started)
        return 0

    ts = _stamp()
    targets = [i for i in items if i.cat == CAT_COPY]
    diffs = [i for i in targets if i.state == DIFF]
    if mode == "check" and diffs:
        out(f"納品先に内容の違う同名ファイルが {len(diffs)} 件あるため、1件もコピーせずに止めました。")
        out("  内容を確かめてから、overwrite（既存のファイルを退避してから上書き）で実行してください。")
        _print_stats(out, started)
        out(f"ログ: {out.save(delivery.work_dir, f'deliver-{mode}', ts)}")
        return 1

    ensure_not_open(path)  # 結果を記録できないのにコピーだけ進まないよう、先に確かめる
    backup_root = delivery.work_dir / "_backup" / f"{ts}_{BACKUP_LABEL}"
    backup_root.mkdir(parents=True, exist_ok=True)
    shutil.copy2(path, backup_root / path.name)

    done = {NEW: 0, SAME: 0, DIFF: 0}
    failed: list[str] = []
    skipped: list[str] = []
    for it in targets:
        rec, dest = it.rec, it.dest
        src = Path(rec.src)
        try:
            if it.state == SAME:
                if rec.st_copy != ST_DONE or rec.copy_hash != it.hash:
                    _mark_done(rec, it.hash, now)
            elif it.state == NEW:
                _mark_done(rec, copy_verified(src, dest), now)
            else:
                if is_editing(dest):
                    skipped.append(f"{_no(rec)} {rec.dest_rel}（納品先が Excel で開かれています）")
                    continue
                saved = backup_file(dest, backup_root, dest_parts(rec))
                rec.backup = str(saved)
                _mark_done(rec, copy_verified(src, dest), now)
            done[it.state] += 1
        except (OSError, CopyError) as e:
            rec.st_copy = ST_NG
            failed.append(f"{_no(rec)} {rec.dest_rel}: {type(e).__name__}: {e}")

    summary, issues = refresh(records, cfg, delivery, [], confirmed, now)
    save_manifest(path, records, issues, summary.rows, confirmed)

    out(f"コピーしました: {NEW} {done[NEW]} 件、上書き（退避してから）{done[DIFF]} 件。"
        f"{SAME}で済にしたもの {done[SAME]} 件")
    if done[DIFF]:
        out(f"上書き前のファイルの退避先: {backup_root}")
    for s in skipped:
        out(f"  飛ばした: {s}")
    if failed:
        out(f"失敗して NG にした行: {len(failed)} 件")
        for f in failed:
            out(f"  {f}")
    out(f"管理表を更新しました: {path}")
    out("\n".join(summary.lines))
    _print_stats(out, started)
    out(f"ログ: {out.save(delivery.work_dir, f'deliver-{mode}', ts)}")
    return 1 if failed else 0


# ---------------------------------------------------------------- restore

def _is_deliver_backup(backup: Path, delivery: Delivery, parts: list[str]) -> bool:
    """backup が <work_dir>/_backup/<ts>_納品コピー/<納品先(相対)> であるか。"""
    try:
        rel = backup.resolve().relative_to((delivery.work_dir / "_backup").resolve())
    except ValueError:
        return False
    return (len(rel.parts) >= 2 and rel.parts[0].endswith(f"_{BACKUP_LABEL}")
            and norm_key("/".join(rel.parts[1:])) == norm_key("/".join(parts)))


def run_restore(cfg, delivery: Delivery, nos: list[int]) -> int:
    started = time.monotonic()
    reset_stats()
    _require_dest(delivery)
    out = Output()
    path = delivery.manifest_path
    records, confirmed = read_manifest(path)
    by_no = {r.no: r for r in records}
    ensure_not_open(path)
    ts = _stamp()
    keep_root = delivery.work_dir / "_backup" / f"{ts}_{RESTORE_LABEL}"
    out(f"納品コピーのバックアップから戻します: {', '.join(f'No.{n}' for n in nos)}")

    todo: list[tuple[Record, Path, list[str], Path]] = []
    refused: list[str] = []
    for n in nos:
        rec = by_no.get(n)
        if rec is None:
            refused.append(f"No.{n}: 管理表にありません")
            continue
        parts = dest_parts(rec)
        dest = dest_path(delivery, rec)
        if not parts or dest is None:
            refused.append(f"No.{n}: 納品先(相対)が空か不正です")
            continue
        if not rec.backup.strip():
            refused.append(f"No.{n}: バックアップがありません（新規でコピーした行は戻せません。"
                           "間違えて置いた場合は、納品先のファイルを手で消して納品コピーを「未」に戻してください）")
            continue
        backup = Path(rec.backup)
        if not _is_deliver_backup(backup, delivery, parts):
            refused.append(f"No.{n}: 納品コピーのバックアップではないため戻しません: {backup}")
            continue
        if not backup.is_file():
            refused.append(f"No.{n}: バックアップのファイルが見つかりません: {backup}")
            continue
        if is_editing(dest):
            refused.append(f"No.{n}: 納品先が Excel で開かれています: {rec.dest_rel}")
            continue
        todo.append((rec, dest, parts, backup))

    failed: list[str] = []
    restored = 0
    if todo:
        keep_root.mkdir(parents=True, exist_ok=True)
        shutil.copy2(path, keep_root / path.name)
    for rec, dest, parts, backup in todo:
        try:
            kept = backup_file(dest, keep_root, parts) if dest.is_file() else None
            copy_verified(backup, dest)
            rec.st_copy, rec.copy_at, rec.copy_hash = ST_TODO, "", ""
            rec.backup = str(kept) if kept else ""
            restored += 1
            out(f"  {_no(rec)} 戻しました: {rec.dest_rel}")
        except (OSError, CopyError) as e:
            failed.append(f"{_no(rec)} {rec.dest_rel}: {type(e).__name__}: {e}")

    for x in refused:
        out(f"  戻さなかった: {x}")
    for x in failed:
        out(f"  失敗: {x}")
    if todo:
        summary, issues = refresh(records, cfg, delivery, [], confirmed, _now())
        save_manifest(path, records, issues, summary.rows, confirmed)
        out(f"{restored} 件を戻し、納品コピーを「未」にしました。戻す前の納品先のファイル: {keep_root}")
        out(f"管理表を更新しました: {path}")
    else:
        out("戻せる行がないため、何も変更していません。")
    _print_stats(out, started)
    out(f"ログ: {out.save(delivery.work_dir, 'restore', ts)}")
    return 1 if (refused or failed) else 0
