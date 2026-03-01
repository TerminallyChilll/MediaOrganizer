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
    if sys.version_info < (3, 9):
        print(f"❌ Error: Python 3.9 or newer is required. You are using Python {sys.version_info.major}.{sys.version_info.minor}.")
        sys.exit(1)

def install_dependencies():
    req_file = os.path.join(os.path.dirname(__file__), 'requirements.txt')
    if not os.path.exists(req_file):
        print(f"⚠️ Warning: {req_file} not found. Skipping auto-install.")
        return

    print("🔍 Checking dependencies...")
    try:
        # First check if we need to install anything
        # using pip freeze or pkg_resources
        import warnings
        with warnings.catch_warnings():
            warnings.simplefilter("ignore", DeprecationWarning)
            import pkg_resources  # type: ignore
        
        with open(req_file, 'r') as f:
            requirements = [line.strip() for line in f if line.strip() and not line.startswith('#')]
            
        pkg_resources.require(requirements)
        print("✅ All dependencies are already installed.")
        
    except Exception:
        print("📦 Missing dependencies detected. Installing now...")
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
                sys.exit(1)

def main():
    check_python_version()
    install_dependencies()
    
    # Now import the actual application and run it
    print("🚀 Launching Media Organizer...\n")
    try:
        import media_organizer  # type: ignore
        media_organizer.main()
    except ImportError as e:
        print(f"❌ Critical Error: Could not load media_organizer.py. Make sure it's in the same directory. ({e})")
        sys.exit(1)
    except Exception as e:
        print(f"❌ Critical Error: {e}")
        import traceback
        traceback.print_exc()
        sys.exit(1)

if __name__ == "__main__":
    main()
