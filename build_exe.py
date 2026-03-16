"""
Build script to create standalone .exe for the RDL application.

Usage:
    pip install pyinstaller
    python build_exe.py

The output will be in the dist/RDL/ folder.
"""
import PyInstaller.__main__
import os
import sys
import shutil

APP_NAME = "RDL"
ROOT_DIR = os.path.dirname(os.path.abspath(__file__))
DIST_DIR = os.path.join(ROOT_DIR, "dist")
BUILD_DIR = os.path.join(ROOT_DIR, "build")


def find_streamlit_path():
    """Find the installed streamlit package path."""
    import streamlit
    return os.path.dirname(streamlit.__file__)


def build():
    streamlit_path = find_streamlit_path()

    # Ensure uploads directory exists for bundling
    uploads_dir = os.path.join(ROOT_DIR, "backend", "uploads")
    os.makedirs(uploads_dir, exist_ok=True)

    # PyInstaller arguments
    args = [
        os.path.join(ROOT_DIR, "launcher.py"),
        "--name", APP_NAME,
        "--noconfirm",
        "--clean",
        # Use onedir mode (more reliable for Streamlit)
        "--onedir",
        # Add backend as data
        "--add-data", f"{os.path.join(ROOT_DIR, 'backend')}{os.pathsep}backend",
        # Add streamlit package (needed for its static files and config)
        "--add-data", f"{streamlit_path}{os.pathsep}streamlit",
        # Hidden imports that PyInstaller may miss
        "--hidden-import", "streamlit",
        "--hidden-import", "streamlit.runtime.scriptrunner",
        "--hidden-import", "streamlit.web.cli",
        "--hidden-import", "pandas",
        "--hidden-import", "openpyxl",
        "--hidden-import", "rapidfuzz",
        "--hidden-import", "altair",
        "--hidden-import", "pyarrow",
        "--hidden-import", "pkg_resources.extern",
        # Console mode so users can see logs
        "--console",
        # Optional: add an icon if you have one
        # "--icon", os.path.join(ROOT_DIR, "icon.ico"),
    ]

    print(f"Building {APP_NAME}.exe ...")
    print(f"Streamlit path: {streamlit_path}")
    print()

    PyInstaller.__main__.run(args)

    # Post-build: create uploads folder in dist
    dist_uploads = os.path.join(DIST_DIR, APP_NAME, "backend", "uploads")
    os.makedirs(dist_uploads, exist_ok=True)

    print()
    print("=" * 60)
    print(f"BUILD COMPLETE!")
    print(f"Output: {os.path.join(DIST_DIR, APP_NAME)}")
    print(f"Run:    {os.path.join(DIST_DIR, APP_NAME, APP_NAME + '.exe')}")
    print("=" * 60)


if __name__ == "__main__":
    build()
