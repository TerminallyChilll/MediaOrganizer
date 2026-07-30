import json
from pathlib import Path

import pytest

from mediaorg.execute import execute, last_run_ops, undo_last_run
from mediaorg.plan import Op, Plan, check_collisions


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
    # An intent is written before the mutation so a crash mid-move is
    # recoverable; completions are what undo replays.
    assert [l["op"] for l in lines if l["op"] == "intent"] == ["intent"]
    completed = [l for l in lines if l["op"] != "intent"]
    kinds = [l["op"] for l in completed]
    # Implicitly-created parent dirs are journaled so undo can remove them.
    assert kinds == ["begin_run", "mkdir", "move", "end_run"]
    assert completed[1]["dst"] == str(tmp_path / "out")
    assert completed[2]["src"] == str(a)
    assert completed[2]["dst"] == str(tmp_path / "out" / "a.mkv")
    # The journal records the file's identity so undo can spot replacement.
    assert completed[2]["size"] == len("a.mkv")


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


# --- Reversibility hardening -------------------------------------------------

import os
import pytest
from pathlib import PureWindowsPath

from mediaorg import execute as ex


@pytest.fixture(autouse=True)
def _no_retry_sleeps(monkeypatch):
    """Keep PermissionError backoff from slowing the suite down."""
    monkeypatch.setattr(ex, "RETRY_DELAYS", ())


def test_all_ops_failing_does_not_lock_out_earlier_undo(tmp_path, journal):
    """A run where every op fails used to make all earlier runs unreachable."""
    a = touch(tmp_path / "a.mkv")
    assert execute(Plan(ops=[Op("move", a, tmp_path / "b.mkv")]), journal).ok

    # Second run: the only op fails, so the run is journaled with zero entries.
    doomed = Plan(ops=[Op("move", tmp_path / "nope.mkv", tmp_path / "c.mkv")])
    result = execute(doomed, journal)
    assert not result.ok and not result.done
    assert [len(r["ops"]) for r in ex.list_runs(journal)] == [1, 0]

    # The empty run is retired and the real one is still reversible.
    assert undo_last_run(journal).ok
    assert a.exists() and not (tmp_path / "b.mkv").exists()
    assert ex.pending_runs(journal) == []


def test_missing_dirs_terminates_on_self_parent():
    """A Windows filesystem root is its own parent — the loop must still end."""
    assert PureWindowsPath("Z:/").parent == PureWindowsPath("Z:/")
    # Real check: a deep non-existent path terminates and is ordered shallowest
    # first, rather than spinning forever.
    got = ex._missing_dirs(Path(os.path.abspath(os.sep)) / "no" / "such" / "dir")
    assert [p.name for p in got] == ["no", "such", "dir"]


def test_hardlink_is_not_treated_as_a_case_only_rename(tmp_path, journal):
    """samefile() is true for hardlinks; the two-step rename must not fire."""
    a = touch(tmp_path / "a.mkv")
    b = tmp_path / "b.mkv"
    try:
        os.link(a, b)
    except (OSError, AttributeError, NotImplementedError):
        pytest.skip("filesystem does not support hardlinks")

    assert not ex._is_case_only_rename(a, b)
    # The planner must refuse it too, rather than waving it through on samefile.
    plan = check_collisions([Op("move", a, b)])
    assert plan.ops == [] and "already exists" in plan.skipped[0][1]
    assert a.exists() and b.exists()


def test_deep_mkdir_is_fully_undone(tmp_path, journal):
    """mkdir used to journal only the leaf, leaving intermediates behind."""
    deep = tmp_path / "x" / "y" / "z"
    assert execute(Plan(ops=[Op("mkdir", None, deep)]), journal).ok
    assert deep.is_dir()
    assert undo_last_run(journal).ok
    assert not (tmp_path / "x").exists()


def test_junk_only_dir_is_quarantined_so_rmdir_and_undo_both_work(tmp_path, journal):
    """.DS_Store used to make every rmdir — and therefore undo — fail."""
    lib = tmp_path / "lib"
    d = lib / "S02"
    d.mkdir(parents=True)
    (d / ".DS_Store").write_text("finder junk")
    before = snapshot(lib)

    result = execute(Plan(ops=[Op("rmdir", None, d)]), journal, roots=[lib])
    assert result.ok, result.failed
    assert not d.exists()
    assert (lib / ".mediaorg_trash" / ".DS_Store").read_text() == "finder junk"

    assert undo_last_run(journal).ok
    assert snapshot(lib) == before


def test_undo_refuses_when_the_file_was_replaced(tmp_path, journal):
    a = touch(tmp_path / "a.mkv")
    moved = tmp_path / "moved.mkv"
    assert execute(Plan(ops=[Op("move", a, moved)]), journal).ok

    moved.write_text("something else entirely, much longer than before")
    result = undo_last_run(journal)
    assert not result.ok
    assert "replaced" in result.failed[0][1] or "modified" in result.failed[0][1]
    assert moved.exists() and not a.exists()

    # --force overrides.
    assert undo_last_run(journal, force=True).ok
    assert a.exists()


def test_recover_removes_a_partial_copy(tmp_path, journal):
    """An intent with no completion means the destination is a partial copy."""
    src = touch(tmp_path / "src.mkv")
    dst = tmp_path / "dst.mkv"
    dst.write_text("half")  # simulate a copy that died partway
    with open(journal, "w", encoding="utf-8") as f:
        f.write(json.dumps({"op": "begin_run", "id": "r1", "v": 2}) + "\n")
        f.write(json.dumps({"op": "intent", "kind": "move", "seq": 1,
                            "src": str(src), "dst": str(dst)}) + "\n")

    notes = ex.recover(journal, dry_run=True)
    assert notes and "incomplete copy" in notes[0]
    assert dst.exists()  # dry run changed nothing

    ex.recover(journal)
    assert not dst.exists() and src.exists()


def test_recover_restores_a_stranded_tmp_file(tmp_path, journal):
    """A crash between the two renames of a case-only rename leaves an orphan."""
    src = tmp_path / "Show S01E01.mkv"
    tmp = tmp_path / ("Show S01E01.mkv" + ex.TMP_SUFFIX)
    tmp.write_text("episode")
    with open(journal, "w", encoding="utf-8") as f:
        f.write(json.dumps({"op": "begin_run", "id": "r1", "v": 2}) + "\n")
        f.write(json.dumps({"op": "intent", "kind": "move", "seq": 1,
                            "src": str(src), "dst": str(tmp_path / "show s01e01.mkv"),
                            "tmp": str(tmp)}) + "\n")

    ex.recover(journal)
    assert src.read_text() == "episode" and not tmp.exists()


def test_undo_run_refuses_out_of_order_when_paths_overlap(tmp_path, journal):
    a = touch(tmp_path / "a.mkv")
    execute(Plan(ops=[Op("move", a, tmp_path / "b.mkv")]), journal)
    execute(Plan(ops=[Op("move", tmp_path / "b.mkv", tmp_path / "c.mkv")]), journal)
    first = ex.list_runs(journal)[0]["id"]

    result, err = ex.undo_run(journal, first)
    assert result is None and "newer" in err
    # Forcing it is allowed, and undoing the newer run first is the clean path.
    assert undo_last_run(journal).ok
    result, err = ex.undo_run(journal, first)
    assert err == "" and result.ok
    assert a.exists()


def test_undo_session_reverses_every_run_of_one_action(tmp_path, journal):
    lib = tmp_path / "lib"
    lib.mkdir()
    a = touch(lib / "a.mkv")
    d = lib / "d"
    before = snapshot(lib)
    execute(Plan(ops=[Op("move", a, lib / "b.mkv")]), journal,
            session="sess1", label="organize")
    execute(Plan(ops=[Op("mkdir", None, d)]), journal,
            session="sess1", label="rename")
    assert len({r["session"] for r in ex.list_runs(journal)}) == 1

    result = ex.undo_session(journal)
    assert result.ok
    assert snapshot(lib) == before
    assert ex.pending_runs(journal) == []


def test_v1_journal_without_new_fields_still_undoes(tmp_path, journal):
    """Existing users have v1 journals on disk; they must keep working."""
    src, dst = tmp_path / "old.mkv", touch(tmp_path / "new.mkv")
    with open(journal, "w", encoding="utf-8") as f:
        f.write(json.dumps({"op": "begin_run", "id": "old1"}) + "\n")
        f.write(json.dumps({"op": "move", "src": str(src), "dst": str(dst)}) + "\n")
        f.write(json.dumps({"op": "end_run", "id": "old1"}) + "\n")

    assert len(last_run_ops(journal)) == 1
    assert undo_last_run(journal).ok
    assert src.exists() and not dst.exists()


def test_v1_undone_run_marker_without_id_is_honoured(tmp_path, journal):
    a = touch(tmp_path / "a.mkv")
    execute(Plan(ops=[Op("move", a, tmp_path / "b.mkv")]), journal)
    with open(journal, "a", encoding="utf-8") as f:
        f.write(json.dumps({"op": "undone_run", "ts": 0}) + "\n")  # v1 form
    assert ex.pending_runs(journal) == []


def test_path_escaping_the_root_is_rejected(tmp_path, journal):
    outside = tmp_path / "outside"
    outside.mkdir()
    inside = tmp_path / "lib"
    inside.mkdir()
    a = touch(inside / "a.mkv")

    result = execute(Plan(ops=[Op("move", a, outside / "a.mkv")]), journal,
                     roots=[inside])
    assert not result.ok and "escapes the library root" in result.failed[0][1]
    assert a.exists()


def test_journal_path_prefers_env_override(monkeypatch, tmp_path):
    target = tmp_path / "custom.jsonl"
    monkeypatch.setenv("MEDIAORG_JOURNAL", str(target))
    assert ex.journal_path() == target


def test_journal_path_is_not_cwd_relative(monkeypatch, tmp_path):
    monkeypatch.delenv("MEDIAORG_JOURNAL", raising=False)
    monkeypatch.chdir(tmp_path)
    resolved = ex.journal_path()
    # Either the app-dir journal, or an adopted pre-existing one — never a
    # brand-new file in an unrelated working directory.
    assert resolved != tmp_path / ex.JOURNAL_FILE


def test_folder_rename_is_undoable_after_its_contents_change(tmp_path, journal):
    """A directory's mtime changes when files inside it are renamed later in
    the same run, so the identity guard must not apply to directories."""
    lib = tmp_path / "lib"
    season = lib / "Season 01"
    season.mkdir(parents=True)
    ep = touch(season / "ep.mkv")
    before = snapshot(lib)

    plan = Plan(ops=[
        Op("move", season, lib / "Season 1"),
        Op("move", lib / "Season 1" / "ep.mkv", lib / "Season 1" / "Ep 1.mkv"),
    ])
    assert execute(plan, journal, roots=[lib]).ok

    result = undo_last_run(journal)
    assert result.ok, result.failed
    assert snapshot(lib) == before
