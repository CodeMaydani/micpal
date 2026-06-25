"""
launch.py
Cross-platform launcher for the מיכפל Streamlit app.

Double-clicking a shortcut to this file (Windows) or running it (Linux) will:
  1. pick the configured port,
  2. if a server is already listening there, just open the browser to it,
  3. otherwise start `streamlit run app.py` headless and open the browser.

This is the single entry point the installers/wizards wrap, so the launch
behavior lives in one place rather than in OS-specific shell scripts.
"""

import os
import socket
import subprocess
import sys
import time
import webbrowser

PORT = int(os.environ.get("MICPAL_PORT", "8501"))
HOST = "localhost"
APP_DIR = os.path.dirname(os.path.abspath(__file__))
APP_FILE = os.path.join(APP_DIR, "app.py")


def _port_in_use(host, port):
    """True if something is already listening on host:port."""
    with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as s:
        s.settimeout(0.5)
        return s.connect_ex((host, port)) == 0


def _wait_until_up(host, port, timeout=30.0):
    """Block until the server accepts connections, or timeout elapses."""
    deadline = time.time() + timeout
    while time.time() < deadline:
        if _port_in_use(host, port):
            return True
        time.sleep(0.3)
    return False


def main():
    url = f"http://{HOST}:{PORT}"

    # Already running -> don't start a second server, just focus the UI.
    if _port_in_use(HOST, PORT):
        webbrowser.open(url)
        return 0

    # Start Streamlit headless (we open the browser ourselves, below).
    cmd = [
        sys.executable,
        "-m",
        "streamlit",
        "run",
        APP_FILE,
        "--server.headless=true",
        f"--server.port={PORT}",
        f"--server.address={HOST}",
    ]
    proc = subprocess.Popen(cmd, cwd=APP_DIR)

    if _wait_until_up(HOST, PORT):
        webbrowser.open(url)

    # Keep this launcher tied to the server process so closing the launcher
    # window stops the app (and so the process isn't orphaned).
    try:
        proc.wait()
    except KeyboardInterrupt:
        proc.terminate()
    return proc.returncode or 0


if __name__ == "__main__":
    sys.exit(main())
