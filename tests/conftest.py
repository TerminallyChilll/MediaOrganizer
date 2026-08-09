"""Shared test fixtures.

Also puts the repo root on sys.path so a bare ``pytest`` works, not just
``python -m pytest`` (there is no packaging metadata to install from).
"""

import sys
from pathlib import Path

import pytest

REPO_ROOT = Path(__file__).resolve().parent.parent
if str(REPO_ROOT) not in sys.path:
    sys.path.insert(0, str(REPO_ROOT))


@pytest.fixture(autouse=True)
def _no_update_check(monkeypatch):
    """Keep the whole suite offline and out of the working tree.

    ``run_wizard`` kicks off a background update check on entry, which shells
    out to git and writes ``.mediaorg_update_check.json`` next to the app —
    outside ``tmp_path``, in the repo. Individual tests remembering a fixture
    is not enough: the one that forgets leaves a stray file behind and burns
    the launch budget. Off by default, everywhere.

    Note this does not cover tests/test_update.py, which drives the caching
    machinery directly and so writes the real cache file on purpose; that is
    pre-existing and wants a `cache_path()` override to fix properly.
    """
    monkeypatch.setenv("MEDIAORG_NO_UPDATE_CHECK", "1")


@pytest.fixture
def journal(tmp_path):
    return tmp_path / "journal.jsonl"


@pytest.fixture
def touch():
    """Create a file (and its parents) whose content is its own name."""
    def _touch(path):
        path = Path(path)
        path.parent.mkdir(parents=True, exist_ok=True)
        path.write_text(path.name)
        return path
    return _touch


@pytest.fixture
def snapshot():
    """Map every path under root to its text (None for directories)."""
    def _snapshot(root: Path) -> dict:
        root = Path(root)
        return {str(p.relative_to(root)): (p.read_text() if p.is_file() else None)
                for p in sorted(root.rglob("*"))}
    return _snapshot


def _probe_case_sensitive(directory: Path) -> bool:
    probe = directory / "MediaOrgCaseProbe"
    probe.mkdir(exist_ok=True)
    try:
        return not (directory / "mediaorgcaseprobe").exists()
    finally:
        probe.rmdir()


@pytest.fixture
def case_sensitive_fs(tmp_path) -> bool:
    """Is the filesystem backing tmp_path case-sensitive?"""
    return _probe_case_sensitive(tmp_path)


@pytest.fixture
def require_case_sensitive_fs(case_sensitive_fs):
    if not case_sensitive_fs:
        pytest.skip("requires a case-sensitive filesystem")


@pytest.fixture
def require_case_insensitive_fs(case_sensitive_fs):
    if case_sensitive_fs:
        pytest.skip("requires a case-insensitive filesystem")
