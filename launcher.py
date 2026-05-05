"""
Xplor Data Migration Validator — Application Launcher
======================================================
This is the entry point when running the packaged application (.exe on Windows,
.app on Mac).

It starts the Streamlit web server and automatically opens the browser so that
the user never needs to touch the command line.

HOW IT WORKS
------------
1. Finds the first available network port (starting at 8501).
2. Spawns a background thread that waits 5 seconds then opens the browser.
3. Starts the Streamlit server in-process (blocking until the window is closed).
4. When the browser tab/window is closed, the server continues running in the
   background — the user must close the terminal/app window to stop it fully.

PACKAGING
---------
Build for Windows (run on a Windows machine):
    pyinstaller xplor_validator.spec

Build for Mac (run on a Mac):
    pyinstaller xplor_validator.spec

The .spec file handles dependency collection for both platforms automatically.
"""

import os
import sys
import socket
import threading
import time
import webbrowser


def find_free_port(start: int = 8501) -> int:
    """
    Finds the first available TCP port starting from `start`.
    Tries up to 100 consecutive ports before giving up.
    """
    for port in range(start, start + 100):
        with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as s:
            try:
                s.bind(("127.0.0.1", port))
                return port
            except OSError:
                continue
    # If all ports are busy, fall back to the default — Streamlit will report
    # a clear error to the user.
    return start


def get_app_path() -> str:
    """
    Returns the absolute path to app.py.

    When running inside a PyInstaller bundle, data files are extracted to a
    temporary directory referenced by sys._MEIPASS.
    When running directly (e.g. `python launcher.py`), files sit alongside
    this script.
    """
    if hasattr(sys, "_MEIPASS"):
        # Running from a PyInstaller bundle
        return os.path.join(sys._MEIPASS, "app.py")
    # Running from source
    return os.path.join(os.path.dirname(os.path.abspath(__file__)), "app.py")


def open_browser(url: str, delay: float = 6.0) -> None:
    """
    Opens the default browser at `url` after a short delay.
    The delay allows the Streamlit server time to start accepting connections.
    Runs in a daemon thread so it does not block the main process.
    """
    time.sleep(delay)
    webbrowser.open(url)


def main() -> None:
    port     = find_free_port()
    url      = f"http://localhost:{port}"
    app_path = get_app_path()

    # Validate that app.py was found — gives a clear error if the bundle is broken
    if not os.path.exists(app_path):
        print(f"ERROR: Could not find app.py at: {app_path}")
        print("The application bundle may be corrupted. Please re-download.")
        input("Press Enter to exit…")
        sys.exit(1)

    # Open the browser in a background thread after the server has started
    threading.Thread(target=open_browser, args=(url,), daemon=True).start()

    print("=" * 60)
    print("  Xplor Data Migration Validator")
    print("=" * 60)
    print(f"  Starting server on {url}")
    print(f"  Your browser will open automatically in a few seconds.")
    print(f"  To stop the application, close this window.")
    print("=" * 60)

    # Run Streamlit in-process.
    # This is the correct approach for PyInstaller — using subprocess would
    # require a separate Python interpreter which is not bundled.
    from streamlit.web import cli as stcli

    sys.argv = [
        "streamlit", "run", app_path,
        f"--server.port={port}",
        "--server.headless=true",
        "--browser.gatherUsageStats=false",
        "--global.developmentMode=false",
        "--theme.base=light",
        "--theme.primaryColor=#0f3460",
        "--theme.backgroundColor=#ffffff",
        "--theme.secondaryBackgroundColor=#f8fafc",
        "--theme.textColor=#1e293b",
    ]

    sys.exit(stcli.main())


if __name__ == "__main__":
    main()
