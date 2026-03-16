"""
Launcher script for packaging the RDL Streamlit app as a standalone .exe.
This script starts the Streamlit server and opens the browser automatically.
"""
import sys
import os
import subprocess
import socket
import webbrowser
import threading
import time


def get_free_port():
    """Find an available port."""
    with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as s:
        s.bind(("", 0))
        return s.getsockname()[1]


def get_app_path():
    """Get the correct path whether running as script or frozen exe."""
    if getattr(sys, "frozen", False):
        return os.path.dirname(sys.executable)
    return os.path.dirname(os.path.abspath(__file__))


def open_browser(port):
    """Open browser after a short delay to let the server start."""
    time.sleep(3)
    webbrowser.open(f"http://localhost:{port}")


def main():
    app_dir = get_app_path()
    main_py = os.path.join(app_dir, "backend", "main.py")

    if not os.path.exists(main_py):
        print(f"ERROR: Cannot find {main_py}")
        input("Press Enter to exit...")
        sys.exit(1)

    port = get_free_port()

    # Open browser in background thread
    threading.Thread(target=open_browser, args=(port,), daemon=True).start()

    print(f"Starting RDL application on http://localhost:{port}")
    print("Close this window to stop the application.\n")

    # Launch streamlit
    cmd = [
        sys.executable, "-m", "streamlit", "run", main_py,
        "--server.port", str(port),
        "--server.headless", "true",
        "--browser.gatherUsageStats", "false",
        "--server.address", "localhost",
    ]

    try:
        process = subprocess.run(cmd, cwd=app_dir)
        sys.exit(process.returncode)
    except KeyboardInterrupt:
        print("\nShutting down...")
        sys.exit(0)


if __name__ == "__main__":
    main()
