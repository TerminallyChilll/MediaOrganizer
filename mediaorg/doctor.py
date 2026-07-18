"""Doctor: diagnose and self-heal common Media Organizer failures.

Run with ``python run.py --doctor`` or ``python -m mediaorg.doctor``.
Every check returns a (status, message) tuple; ``fix()`` attempts repair.
"""

import json
import os
import subprocess
import sys
from pathlib import Path

# ── helpers ──────────────────────────────────────────────────────────────

def _ok(msg: str) -> tuple[str, str]:   return "OK", msg
def _warn(msg: str) -> tuple[str, str]: return "WARN", msg
def _err(msg: str) -> tuple[str, str]:  return "FAIL", msg


def _pkg_dir() -> Path:
    """Absolute path to the mediaorg package root (where run.py lives)."""
    return Path(__file__).resolve().parent.parent


# ── checks ───────────────────────────────────────────────────────────────

def check_python() -> tuple[str, str]:
    """Python >= 3.11?"""
    vi = sys.version_info
    if vi < (3, 11):
        return _err(
            f"Python {vi.major}.{vi.minor}.{vi.micro} — need 3.11+.\n"
            f"  Download from https://www.python.org/downloads/"
        )
    return _ok(f"Python {vi.major}.{vi.minor}.{vi.micro}")


def check_pip() -> tuple[str, str]:
    """Is pip importable?"""
    try:
        import pip  # noqa: F401
        return _ok("pip is available")
    except ImportError:
        return _err("pip module not found — Python was installed without pip")


def check_deps() -> tuple[str, str]:
    """Are the four runtime deps importable?"""
    missing = []
    for pkg in ("pandas", "openpyxl", "tqdm", "guessit"):
        try:
            __import__(pkg)
        except ImportError:
            missing.append(pkg)
    if missing:
        return _err(f"Missing packages: {', '.join(missing)}")
    return _ok("All 4 runtime dependencies present")


def check_package_structure() -> tuple[str, str]:
    """Does the mediaorg/ package look intact?"""
    root = _pkg_dir()
    required = [
        root / "run.py",
        root / "requirements.txt",
        root / "mediaorg" / "__init__.py",
        root / "mediaorg" / "wizard.py",
        root / "mediaorg" / "plan.py",
        root / "mediaorg" / "execute.py",
        root / "mediaorg" / "scan.py",
        root / "mediaorg" / "excel.py",
        root / "mediaorg" / "parse.py",
        root / "mediaorg" / "llm.py",
        root / "mediaorg" / "extfix.py",
    ]
    missing = [str(p.relative_to(root)) for p in required if not p.exists()]
    if missing:
        return _err(f"Missing files: {', '.join(missing)}")
    return _ok("Package structure intact")


def _validate_json(path: Path) -> tuple[str, str]:
    if not path.exists():
        return _warn(f"{path.name} not present (will be created on first use)")
    try:
        data = json.loads(path.read_text(encoding="utf-8"))
        if not isinstance(data, (dict, list)):
            return _err(f"{path.name} is not a JSON object/array")
        return _ok(f"{path.name} — valid JSON")
    except json.JSONDecodeError as e:
        return _err(f"{path.name} is corrupted: {e}")


def check_configs() -> list[tuple[str, str]]:
    """Validate every JSON state file in the working directory."""
    results = []
    for name in (".media_renamer_config.json", ".media_llm_config.json",
                  "custom_strip_patterns.json", "mediaorg_journal.jsonl"):
        results.append(_validate_json(Path(name)))
    return results


def check_journal() -> tuple[str, str]:
    """Does the journal have torn (unfinished) runs?"""
    jp = Path("mediaorg_journal.jsonl")
    if not jp.exists():
        return _ok("No journal file — nothing to check")
    entries = []
    with open(jp, encoding="utf-8") as f:
        for line in f:
            line = line.strip()
            if not line:
                continue
            try:
                entries.append(json.loads(line))
            except json.JSONDecodeError:
                return _warn("Journal has unparseable lines (torn writes)")
    depth = 0
    for e in entries:
        if e.get("op") == "begin_run":
            depth += 1
        elif e.get("op") == "end_run":
            depth -= 1
        elif e.get("op") == "undone_run":
            pass  # legitimate
    if depth > 0:
        return _warn(f"Journal has {depth} unfinished run(s) — may hide older undo-able runs")
    elif depth < 0:
        return _warn("Journal has unmatched end_run(s) — may be harmless")
    return _ok("Journal runs are balanced (no torn runs)")


def check_stdout() -> tuple[str, str]:
    """Can we output UTF-8?"""
    for stream in (sys.stdout, sys.stderr):
        if hasattr(stream, "reconfigure"):
            try:
                stream.reconfigure(encoding="utf-8", errors="replace")
            except Exception:
                pass
    try:
        print("\u2705", end="", flush=True)
        print("\r", end="", flush=True)
        return _ok("Unicode output works")
    except UnicodeEncodeError:
        return _warn("Terminal does not support Unicode — UI markers may look garbled")


# ── fixers ────────────────────────────────────────────────────────────────

def fix_pip() -> tuple[str, str]:
    """Bootstrap pip via ensurepip."""
    try:
        subprocess.check_call(
            [sys.executable, "-m", "ensurepip", "--upgrade"],
            stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL,
        )
        return _ok("pip bootstrapped via ensurepip")
    except subprocess.CalledProcessError as e:
        return _err(f"ensurepip failed (exit {e.returncode}).\n"
                     f"  Reinstall Python from https://www.python.org/downloads/\n"
                     f"  and make sure 'pip' is checked in the installer.")


def fix_deps() -> tuple[str, str]:
    """pip install -r requirements.txt."""
    req = _pkg_dir() / "requirements.txt"
    if not req.exists():
        return _err("requirements.txt missing — cannot install")
    try:
        subprocess.check_call(
            [sys.executable, "-m", "pip", "install", "-r", str(req), "--quiet"],
            stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL,
        )
        return _ok("Dependencies installed")
    except subprocess.CalledProcessError:
        try:
            subprocess.check_call(
                [sys.executable, "-m", "pip", "install", "--user", "-r", str(req), "--quiet"],
                stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL,
            )
            return _ok("Dependencies installed (--user)")
        except subprocess.CalledProcessError as e:
            return _err(f"pip install failed (exit {e.returncode}).\n"
                         f"  Run manually: pip install -r requirements.txt")


def fix_corrupted_config(path: Path) -> tuple[str, str]:
    """Back up and remove a corrupted JSON file so it can be recreated."""
    backup = path.with_suffix(path.suffix + ".bak")
    try:
        path.rename(backup)
        return _ok(f"{path.name} backed up to {backup.name} — will be recreated on next use")
    except OSError as e:
        return _err(f"Could not back up {path.name}: {e}")


def fix_journal_torn() -> tuple[str, str]:
    """Close any open begin_run with a synthetic end_run so older runs are
    reachable via Undo."""
    jp = Path("mediaorg_journal.jsonl")
    if not jp.exists():
        return _ok("No journal file")
    entries = []
    with open(jp, encoding="utf-8") as f:
        for line in f:
            line = line.strip()
            if not line:
                continue
            try:
                entries.append(json.loads(line))
            except json.JSONDecodeError:
                continue
    depth = 0
    for e in entries:
        if e.get("op") == "begin_run":
            depth += 1
        elif e.get("op") == "end_run":
            depth -= 1
    if depth <= 0:
        return _ok("Journal already balanced")
    import time as _time
    with open(jp, "a", encoding="utf-8") as jf:
        for _ in range(depth):
            jf.write(json.dumps({"op": "end_run", "id": "doctor-fix",
                                 "ts": _time.time()}) + "\n")
    return _ok(f"Closed {depth} torn run(s) with synthetic end_run markers")


# ── runner ────────────────────────────────────────────────────────────────

def run_doctor(*, auto_fix: bool = False) -> int:
    """Run all checks, optionally auto-fixing failures. Returns exit code."""
    print("=" * 60)
    print("Media Organizer — Doctor")
    print("=" * 60)

    issues = 0
    warnings = 0
    fixed = 0

    def report(status: str, label: str, detail: str) -> None:
        nonlocal issues, warnings
        icon = {"OK": "✅", "WARN": "⚠️", "FAIL": "❌"}.get(status, "?")
        print(f"\n{icon} {label}")
        for line in detail.split("\n"):
            print(f"   {line}")
        if status == "FAIL":
            issues += 1
        elif status == "WARN":
            warnings += 1

    # 1. Python version
    s, msg = check_python()
    report(s, "Python version", msg)
    if s == "FAIL":
        print("\n❌ Doctor cannot continue — upgrade Python first.")
        return 1

    # 2. Pip
    s, msg = check_pip()
    report(s, "pip", msg)
    if s == "FAIL" and auto_fix:
        s2, m2 = fix_pip()
        report(s2, "pip (fix)", m2)
        if s2 == "OK":
            fixed += 1
            issues -= 1

    # 3. Dependencies (skip if pip still broken)
    s2, _ = check_pip()
    if s2 == "OK":
        s, msg = check_deps()
        report(s, "Dependencies", msg)
        if s == "FAIL" and auto_fix:
            s3, m3 = fix_deps()
            report(s3, "Dependencies (fix)", m3)
            if s3 == "OK":
                # re-check
                s, msg = check_deps()
                if s == "OK":
                    fixed += 1

    # 4. Package structure
    s, msg = check_package_structure()
    report(s, "Package files", msg)

    # 5. Config files
    print("\n── Config files ──")
    for s, msg in check_configs():
        name = msg.split(" ")[0]
        report(s, name, msg)
        if s == "FAIL" and auto_fix:
            fname = msg.split(" ")[0]
            sf, mf = fix_corrupted_config(Path(fname))
            report(sf, f"{fname} (fix)", mf)
            if sf == "OK":
                fixed += 1

    # 6. Journal integrity
    s, msg = check_journal()
    report(s, "Journal", msg)
    if s == "WARN" and auto_fix and "unfinished" in msg:
        sf, mf = fix_journal_torn()
        report(sf, "Journal (fix)", mf)
        if sf == "OK":
            fixed += 1

    # 7. Unicode output
    s, msg = check_stdout()
    report(s, "Terminal Unicode", msg)

    # Summary
    print("\n" + "=" * 60)
    if issues == 0:
        if warnings:
            print(f"✅ All checks passed ({warnings} minor warning(s)).")
        else:
            print("✅ All checks passed — ready to run.")
        return 0
    else:
        print(f"❌ {issues} issue(s) found, {fixed} fixed, {warnings} warning(s).")
        if not auto_fix:
            print("   Re-run with --doctor --fix to attempt automatic repair.")
        return 0 if issues == 0 else 1


# ── CLI ───────────────────────────────────────────────────────────────────

if __name__ == "__main__":
    auto = "--fix" in sys.argv
    sys.exit(run_doctor(auto_fix=auto))
