"""
Office Automation Suite - Native Desktop Application Launcher
Provides a standalone Windows software experience with dedicated application window,
while running the local data processor and extension receiver in the background.
"""
import sys
import os
import time
import socket
import threading
import urllib.request
import webbrowser

# Protect against PyInstaller console=False NoneType print crashes
class NullWriter:
    def write(self, s): pass
    def flush(self): pass

if sys.stdout is None:
    sys.stdout = NullWriter()
if sys.stderr is None:
    sys.stderr = NullWriter()

DEFAULT_PORT = 5055
HOST = "127.0.0.1"


def is_port_in_use(port):
    """Check if a local TCP port is already open."""
    try:
        with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as s:
            s.settimeout(0.5)
            return s.connect_ex((HOST, port)) == 0
    except Exception:
        return False


def is_server_healthy(url, timeout=1.0):
    """Verify whether our Flask server is active and returning HTTP 200."""
    try:
        req = urllib.request.Request(f"{url}/api/health")
        with urllib.request.urlopen(req, timeout=timeout) as res:
            return res.status == 200
    except Exception:
        return False


def find_available_port(start_port=5055, max_attempts=15):
    """Finds an unused port or one where our application is already running."""
    # 1. Check if our application is already running on any of these ports
    for p in range(start_port, start_port + max_attempts):
        url = f"http://{HOST}:{p}"
        if is_server_healthy(url):
            return p, True  # Already running our app!

    # 2. Find a port that we can actually bind to
    for p in range(start_port, start_port + max_attempts):
        try:
            with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as s:
                s.bind((HOST, p))
                return p, False  # Free port to bind
        except Exception:
            continue

    return start_port, False


def start_flask(port):
    """Starts the Flask server on the designated port."""
    from app import app
    try:
        # Werkzeug server
        app.run(host=HOST, port=port, debug=False, use_reloader=False, threaded=True)
    except Exception as e:
        pass


def wait_for_server(url, timeout=8):
    """Polls until the server responds to /api/health or timeout occurs."""
    start_t = time.time()
    while time.time() - start_t < timeout:
        if is_server_healthy(url, timeout=0.5):
            return True
        time.sleep(0.15)
    return False


def main():
    port, already_running = find_available_port(DEFAULT_PORT)
    app_url = f"http://{HOST}:{port}"

    # 1. Start background server if not already running
    if not already_running:
        t = threading.Thread(target=start_flask, args=(port,), daemon=True)
        t.start()
        wait_for_server(app_url)

    # 2. Try launching native desktop window with pywebview
    try:
        import webview

        class DesktopApi:
            def browse_excel(self, file_type="master"):
                # Use pywebview active window to trigger native file dialog
                try:
                    w = webview.active_window()
                    if not w and webview.windows:
                        w = webview.windows[0]
                    if w:
                        result = w.create_file_dialog(
                            webview.OPEN_DIALOG,
                            allow_multiple=False,
                            file_types=('Excel Spreadsheets (*.xlsx;*.xls;*.xlsm)', 'All files (*.*)')
                        )
                        if result and len(result) > 0:
                            return os.path.normpath(result[0])
                except Exception:
                    pass

                # Fallback to tkinter file dialog if native dialog encounters issues
                try:
                    import tkinter as tk
                    from tkinter import filedialog
                    root = tk.Tk()
                    root.withdraw()
                    root.attributes('-topmost', True)
                    selected = filedialog.askopenfilename(
                        title="Select Excel File",
                        filetypes=[("Excel Spreadsheets", "*.xlsx *.xls *.xlsm"), ("All Files", "*.*")]
                    )
                    root.destroy()
                    if selected:
                        return os.path.normpath(selected)
                except Exception:
                    pass

                return ""

        api = DesktopApi()
        webview.create_window(
            title="Office Automation - MIS Claims & Estimate Parts",
            url=app_url,
            width=1340,
            height=860,
            min_size=(960, 600),
            background_color="#0f172a",
            text_select=True,
            confirm_close=False,
            js_api=api
        )
        webview.start(debug=False)

    except Exception:
        # Fallback to default system browser if webview encounters native runtime issues
        try:
            webbrowser.open(app_url)
        except Exception:
            pass

        # Keep background server alive
        try:
            while True:
                time.sleep(1)
        except KeyboardInterrupt:
            pass

    sys.exit(0)


if __name__ == "__main__":
    main()
