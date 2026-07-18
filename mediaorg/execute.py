"""Execute plans with a crash-safe JSONL journal; undo by reverse replay.

Every SUCCESSFUL filesystem change is appended to the journal (with the
actual paths used) and flushed immediately, so the journal always matches
disk even after a crash or permission error mid-run.
"""

import json
import os
import shutil
import time
import uuid
from dataclasses import dataclass, field
from pathlib import Path

from .plan import Op, Plan

JOURNAL_FILE = "mediaorg_journal.jsonl"


@dataclass
class ExecResult:
    done: list[Op] = field(default_factory=list)
    failed: list[tuple[Op, str]] = field(default_factory=list)
    dry_run: bool = False

    @property
    def ok(self) -> bool:
        return not self.failed


def _same_file(a: Path, b: Path) -> bool:
    try:
        return a.samefile(b)
    except OSError:
        return False


def _missing_parents(dst: Path) -> list[Path]:
    """Ancestors of dst that don't exist yet, shallowest first."""
    missing = []
    p = dst.parent
    while not p.exists():
        missing.append(p)
        p = p.parent
    return list(reversed(missing))


def _safe_path(p: Path) -> Path:
    """Reject path traversal via '..' components and symlink escapes."""
    if '..' in p.parts:
        raise ValueError(f"path traversal rejected: {p}")
    # Reject paths containing symlink components that could redirect
    # outside the intended directory tree.
    for ancestor in (p if p.is_absolute() else Path.cwd() / p).parents:
        try:
            if ancestor.is_symlink():
                raise ValueError(
                    f"symlink traversal rejected: {p} "
                    f"(component {ancestor} is a symlink)")
        except OSError:
            pass  # permissions, deleted file — not our concern
    return p.resolve()


def _do_move(src: Path, dst: Path) -> None:
    if dst.exists() and _same_file(src, dst) and str(src) != str(dst):
        # Case-only rename on a case-insensitive filesystem: two-step.
        tmp = src.with_name(src.name + ".mediaorg_tmp")
        os.rename(src, tmp)
        os.rename(tmp, dst)
        return
    if os.path.lexists(dst):
        raise FileExistsError(f"target already exists: {dst}")
    dst.parent.mkdir(parents=True, exist_ok=True)
    shutil.move(str(src), str(dst))


def execute(plan: Plan, journal: Path, dry_run: bool = False) -> ExecResult:
    """Apply plan ops in order. Journals each successful op immediately."""
    result = ExecResult(dry_run=dry_run)
    if dry_run or not plan.ops:
        result.done = list(plan.ops)
        return result

    run_id = uuid.uuid4().hex[:12]
    ops_logged = 0
    with open(journal, "a", encoding="utf-8") as jf:
        def log(entry: dict) -> None:
            jf.write(json.dumps(entry, ensure_ascii=False) + "\n")
            jf.flush()
            os.fsync(jf.fileno())

        log({"op": "begin_run", "id": run_id, "ts": time.time()})
        for op in plan.ops:
            created: list[Path] = []
            try:
                if op.kind == "move":
                    created = _missing_parents(op.dst)
                    _do_move(_safe_path(op.src), _safe_path(op.dst))
                elif op.kind == "mkdir":
                    _safe_path(op.dst).mkdir(parents=True, exist_ok=True)
                elif op.kind == "rmdir":
                    _safe_path(op.dst).rmdir()  # fails if non-empty — that's the safety
                else:
                    raise ValueError(f"unknown op kind: {op.kind}")
            except (OSError, ValueError) as e:
                result.failed.append((op, str(e)))
                continue
            # Journal implicitly-created parents BEFORE the move entry so
            # reverse replay undoes the move first, then rmdirs the dirs.
            for d in created:
                log({"op": "mkdir", "src": None, "dst": str(d), "ts": time.time()})
            log({"op": op.kind,
                 "src": str(op.src) if op.src else None,
                 "dst": str(op.dst), "ts": time.time()})
            result.done.append(op)
            ops_logged += 1
        log({"op": "end_run", "id": run_id, "ts": time.time()})
    return result


def _read_journal(journal: Path) -> list[dict]:
    entries = []
    with open(journal, encoding="utf-8") as f:
        for line in f:
            line = line.strip()
            if line:
                try:
                    entries.append(json.loads(line))
                except json.JSONDecodeError:
                    continue  # torn write from a crash — ignore
    return entries


def last_run_ops(journal: Path) -> list[dict]:
    """Ops of the most recent run that hasn't been undone yet."""
    if not journal.exists():
        return []
    entries = _read_journal(journal)
    runs: list[list[dict]] = []
    current: list[dict] | None = None
    for e in entries:
        if e["op"] == "begin_run":
            current = []
            runs.append(current)
        elif e["op"] == "undone_run":
            if runs:
                runs.pop()
        elif e["op"] == "end_run":
            current = None
        elif current is not None:
            current.append(e)
        # entries after end_run but before next begin_run can't happen in
        # normal operation; a crash mid-run just leaves the run open.
    return runs[-1] if runs else []


def undo_last_run(journal: Path, dry_run: bool = False) -> ExecResult:
    """Reverse-replay the last un-undone run. Repeatable for earlier runs."""
    result = ExecResult(dry_run=dry_run)
    ops = last_run_ops(journal)
    if not ops:
        return result

    reverse: list[Op] = []
    for e in reversed(ops):
        if e["op"] == "move":
            reverse.append(Op("move", Path(e["dst"]), Path(e["src"])))
        elif e["op"] == "mkdir":
            reverse.append(Op("rmdir", None, Path(e["dst"])))
        elif e["op"] == "rmdir":
            reverse.append(Op("mkdir", None, Path(e["dst"])))

    if dry_run:
        result.done = reverse
        return result

    for op in reverse:
        try:
            if op.kind == "move":
                _do_move(_safe_path(op.src), _safe_path(op.dst))
            elif op.kind == "mkdir":
                _safe_path(op.dst).mkdir(parents=True, exist_ok=True)
            elif op.kind == "rmdir":
                _safe_path(op.dst).rmdir()
        except OSError as e:
            # If a reverse move fails because the source is already gone,
            # it was already reverted in a prior partial undo — count it done.
            if op.kind == "move" and not os.path.lexists(op.src):
                result.done.append(op)
                continue
            result.failed.append((op, str(e)))
            continue
        except ValueError as e:
            result.failed.append((op, str(e)))
            continue
        result.done.append(op)

    if result.ok:
        with open(journal, "a", encoding="utf-8") as jf:
            jf.write(json.dumps({"op": "undone_run", "ts": time.time()}) + "\n")
    # Partial failure: already-reverted ops (source-missing on reverse move)
    # are treated as done so the run can be marked undone. Genuine failures
    # (e.g. permission errors) keep the run alive for retry.
    return result
