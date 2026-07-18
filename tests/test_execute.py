import json
from pathlib import Path

import pytest

from mediaorg.execute import execute, last_run_ops, undo_last_run
from mediaorg.plan import Op, Plan


def touch(path):
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(path.name)
    return path


def snapshot(root: Path) -> dict:
    return {str(p.relative_to(root)): (p.read_text() if p.is_file() else None)
            for p in sorted(root.rglob("*"))}


@pytest.fixture
def journal(tmp_path):
    return tmp_path / "journal.jsonl"


def test_journal_matches_disk(tmp_path, journal):
    a = touch(tmp_path / "a.mkv")
    plan = Plan(ops=[Op("move", a, tmp_path / "out" / "a.mkv")])
    result = execute(plan, journal)
    assert result.ok
    assert (tmp_path / "out" / "a.mkv").exists() and not a.exists()

    lines = [json.loads(l) for l in journal.read_text().splitlines()]
    kinds = [l["op"] for l in lines]
    # Implicitly-created parent dirs are journaled so undo can remove them.
    assert kinds == ["begin_run", "mkdir", "move", "end_run"]
    assert lines[1]["dst"] == str(tmp_path / "out")
    assert lines[2]["src"] == str(a)
    assert lines[2]["dst"] == str(tmp_path / "out" / "a.mkv")


def test_dry_run_touches_nothing(tmp_path, journal):
    a = touch(tmp_path / "a.mkv")
    plan = Plan(ops=[Op("move", a, tmp_path / "out" / "a.mkv")])
    result = execute(plan, journal, dry_run=True)
    assert result.dry_run and result.done == plan.ops
    assert a.exists() and not (tmp_path / "out").exists()
    assert not journal.exists()


def test_undo_restores_tree(tmp_path, journal):
    root = tmp_path / "lib"
    touch(root / "Show.S01E01.mkv")
    touch(root / "Show.S01E01.srt")
    (root / "S02").mkdir()
    touch(root / "S02" / "ep.mkv")
    before = snapshot(root)

    plan = Plan(ops=[
        Op("move", root / "Show.S01E01.mkv", root / "Season 1" / "Show.S01E01.mkv"),
        Op("move", root / "Show.S01E01.srt", root / "Season 1" / "Show.S01E01.srt"),
        Op("move", root / "S02", root / "Season 2"),
    ])
    assert execute(plan, journal).ok
    assert snapshot(root) != before

    result = undo_last_run(journal)
    assert result.ok
    assert snapshot(root) == before
    # Journal marked undone: nothing left to undo.
    assert last_run_ops(journal) == []


def test_undo_undoes_runs_newest_first(tmp_path, journal):
    a = touch(tmp_path / "a.mkv")
    execute(Plan(ops=[Op("move", a, tmp_path / "b.mkv")]), journal)
    execute(Plan(ops=[Op("move", tmp_path / "b.mkv", tmp_path / "c.mkv")]), journal)

    undo_last_run(journal)  # c -> b
    assert (tmp_path / "b.mkv").exists()
    undo_last_run(journal)  # b -> a
    assert (tmp_path / "a.mkv").exists()
    assert last_run_ops(journal) == []


def test_midrun_failure_journals_prior_ops_and_is_undoable(tmp_path, journal):
    a = touch(tmp_path / "a.mkv")
    b = touch(tmp_path / "b.mkv")
    touch(tmp_path / "out" / "b.mkv")  # pre-existing conflict for op 2
    plan = Plan(ops=[
        Op("move", a, tmp_path / "out" / "a.mkv"),
        Op("move", b, tmp_path / "out" / "b.mkv"),
    ])
    result = execute(plan, journal)
    assert len(result.done) == 1 and len(result.failed) == 1
    assert "already exists" in result.failed[0][1]
    # b untouched, a moved and journaled.
    assert b.exists()
    ops = last_run_ops(journal)
    assert len(ops) == 1 and ops[0]["dst"].endswith("a.mkv")

    assert undo_last_run(journal).ok
    assert a.exists()


def test_undo_partial_failure_is_retryable(tmp_path, journal):
    a = touch(tmp_path / "a.mkv")
    execute(Plan(ops=[Op("move", a, tmp_path / "moved.mkv")]), journal)
    # Sabotage: occupy the undo target.
    touch(tmp_path / "a.mkv")
    result = undo_last_run(journal)
    assert not result.ok
    # Run NOT marked undone — still retryable after user clears the conflict.
    assert last_run_ops(journal)
    (tmp_path / "a.mkv").unlink()
    assert undo_last_run(journal).ok
    assert (tmp_path / "a.mkv").read_text() == "a.mkv"


def test_rmdir_only_removes_empty(tmp_path, journal):
    d = tmp_path / "dir"
    d.mkdir()
    touch(d / "keep.txt")
    result = execute(Plan(ops=[Op("rmdir", None, d)]), journal)
    assert not result.ok
    assert d.exists() and (d / "keep.txt").exists()


def test_mkdir_rmdir_roundtrip_via_undo(tmp_path, journal):
    d = tmp_path / "newdir"
    execute(Plan(ops=[Op("mkdir", None, d)]), journal)
    assert d.is_dir()
    undo_last_run(journal)
    assert not d.exists()


def test_case_only_rename(tmp_path, journal):
    src = touch(tmp_path / "show s01e01.mkv")
    dst = tmp_path / "Show S01E01.mkv"
    result = execute(Plan(ops=[Op("move", src, dst)]), journal)
    assert result.ok
    assert dst.name in [p.name for p in tmp_path.iterdir()]
    # Works on both case-sensitive and case-insensitive filesystems.
    undo_last_run(journal)
    assert src.name in [p.name for p in tmp_path.iterdir()]


def test_undo_with_no_journal(tmp_path, journal):
    result = undo_last_run(journal)
    assert result.ok and result.done == []


def test_torn_journal_line_ignored(tmp_path, journal):
    a = touch(tmp_path / "a.mkv")
    execute(Plan(ops=[Op("move", a, tmp_path / "b.mkv")]), journal)
    with open(journal, "a") as f:
        f.write('{"op": "move", "src"')  # simulated crash mid-write
    assert len(last_run_ops(journal)) == 1
    assert undo_last_run(journal).ok
    assert a.exists()
