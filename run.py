#!/usr/bin/env python3
"""
Launcher for Media Organizer.
Automatically verifies and installs missing dependencies from requirements.txt,
then launches the main application.
"""

import sys
import subprocess
import os

# Fix emoji printing in Windows cmd limitataions
if hasattr(sys.stdout, 'reconfigure'):
    try:
        sys.stdout.reconfigure(encoding='utf-8')  # type: ignore
    except Exception:
        pass

def check_python_version():
    if sys.version_info < (3, 11):
        print(f"❌ Error: Python 3.11 or newer is required. You are using Python {sys.version_info.major}.{sys.version_info.minor}.")
        sys.exit(1)

def install_dependencies():
    req_file = os.path.join(os.path.dirname(__file__), 'requirements.txt')
    if not os.path.exists(req_file):
        print(f"⚠️ Warning: {req_file} not found. Skipping auto-install.")
        return

    print("🔍 Checking dependencies...")
    try:
        import pandas, openpyxl, tqdm, guessit  # noqa: F401
        # Verify versions meet minimums from requirements.txt
        from importlib.metadata import version as _pkg_version
        _min_versions = {'pandas': '2.2', 'openpyxl': '3.1', 'tqdm': '4.66', 'guessit': '3.8'}
        _stale = []
        for _pkg, _min in _min_versions.items():
            _inst = _pkg_version(_pkg)
            # Compare as integer tuples — string comparison of version
            # numbers is fragile (e.g. "2.10" < "2.2" lexicographically).
            # Strip trailing non-numeric segments ("2.3a1" → "2.3") so
            # int() doesn't choke on pre-release suffixes; treat an
            # unparseable version as acceptable rather than crashing.
            _inst_num = _inst.split('-')[0].split('+')[0]
            try:
                _inst_tup = tuple(int(x) for x in _inst_num.split('.'))
            except ValueError:
                continue  # can't parse — assume it's fine
            _min_tup = tuple(int(x) for x in _min.split('.'))
            if _inst_tup < _min_tup:
                _stale.append(f"{_pkg}=={_inst} (need >={_min})")
        if _stale:
            print(f"⚠️  Outdated packages: {', '.join(_stale)}")
            raise ImportError("outdated dependencies")
        print("✅ All dependencies are already installed.")
    except ImportError:
        print("📦 Missing dependencies detected. Installing now...")

        # If pip itself is missing, try bootstrapping it first.
        try:
            import pip  # noqa: F401
        except ImportError:
            print("   pip module not found — bootstrapping via ensurepip...")
            try:
                subprocess.check_call([sys.executable, "-m", "ensurepip", "--upgrade"],
                                      stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
                print("   pip bootstrapped.")
            except subprocess.CalledProcessError:
                print("❌ pip is not installed and ensurepip failed.")
                print("   Reinstall Python from https://www.python.org/downloads/")
                print("   and make sure 'pip' is checked in the installer.")
                print("   Or run: python run.py --doctor --fix")
                sys.exit(1)

        try:
            # Install dependencies quietly
            subprocess.check_call([sys.executable, "-m", "pip", "install", "-r", req_file, "--quiet"])
            print("✅ Dependencies installed successfully!\n")
        except subprocess.CalledProcessError:
            print("⚠️ Standard install failed, trying with --user flag (permissions issue)...")
            try:
                subprocess.check_call([sys.executable, "-m", "pip", "install", "--user", "-r", req_file, "--quiet"])
                print("✅ Dependencies installed successfully!\n")
            except subprocess.CalledProcessError as e:
                print(f"❌ Failed to install dependencies: {e}")
                print("Please try running: pip install -r requirements.txt manually.")
                print("Or run: python run.py --doctor --fix")
                sys.exit(1)

def main():
    # Version guard first: doctor.py uses modern syntax and would crash on
    # import under an unsupported Python before showing the upgrade message.
    check_python_version()

    # ── Doctor mode ──
    if "--doctor" in sys.argv:
        from mediaorg.doctor import run_doctor
        auto = "--fix" in sys.argv
        sys.exit(run_doctor(auto_fix=auto))

    # ── Version / update ──
    # Handled before install_dependencies(): pulling a fix is exactly what you
    # want to do when the dependency install is what's broken, and none of
    # these touch pandas/guessit. The parsing lives in mediaorg.update so this
    # and the wizard's parser cannot disagree about what was typed.
    from mediaorg import update
    code = update.dispatch_cli(sys.argv[1:])
    if code is not None:
        sys.exit(code)

    install_dependencies()
    
    # Now import the actual application and run it
    print("🚀 Launching Media Organizer...\n")
    try:
        from mediaorg import wizard  # type: ignore
        wizard.main()
    except ImportError as e:
        print(f"❌ Critical Error: Could not load the mediaorg package. Make sure it's in the same directory. ({e})")
        print("Try running: python run.py --doctor")
        sys.exit(1)
    except Exception as e:
        print(f"❌ Critical Error: {e}")
        import traceback
        traceback.print_exc()
        print("\nTry running: python run.py --doctor")
        sys.exit(1)

if __name__ == "__main__":
    main()
