"""
Flask Application for Insurance Claim MIS Extractor.
Runs on localhost to provide an intuitive interface for PDF data extraction and Excel appending.
"""

import os
import sys
import json
import subprocess
import threading
from datetime import datetime
from flask import Flask, request, jsonify, render_template, send_from_directory
from werkzeug.utils import secure_filename

# Protect against PyInstaller console=False NoneType print crashes
class NullWriter:
    def write(self, s): pass
    def flush(self): pass

if sys.stdout is None:
    sys.stdout = NullWriter()
if sys.stderr is None:
    sys.stderr = NullWriter()

from extractor import extract_data_from_pdf, DEFAULT_TEMPLATES
from excel_handler import get_excel_info, append_rows_to_excel, check_file_locked
from estimate_handler import parse_estimate_pdf, inspect_estimate_excel, write_parts_to_excel

if getattr(sys, "frozen", False):
    BASE_DIR = os.path.dirname(sys.executable)
    TEMPLATE_DIR = os.path.join(getattr(sys, "_MEIPASS", BASE_DIR), "templates")
    app = Flask(__name__, template_folder=TEMPLATE_DIR)
else:
    BASE_DIR = os.path.dirname(os.path.abspath(__file__))
    app = Flask(__name__)

app.config["SECRET_KEY"] = "office-claim-secret-key-2026"
UPLOAD_FOLDER = os.path.join(BASE_DIR, "uploads")
CONFIG_FILE = os.path.join(BASE_DIR, "config.json")
os.makedirs(UPLOAD_FOLDER, exist_ok=True)

# In-memory queue for web claims received via extension or web portal (protected by lock)
web_claims_lock = threading.Lock()
web_claims_queue = []

@app.after_request
def add_cors_headers(response):
    response.headers["Access-Control-Allow-Origin"] = "*"
    response.headers["Access-Control-Allow-Headers"] = "Content-Type, Authorization"
    response.headers["Access-Control-Allow-Methods"] = "GET, POST, OPTIONS, DELETE, PUT"
    return response

DEFAULT_CONFIG = {
    "master_excel_path": "",
    "selected_sheet": "",
    "auto_increment_ref": True,
    "default_survey_date_today": True,
    "create_backup": True,
    "estimate_excel_path": "",
    "estimate_selected_sheet": "",
    "estimate_start_row": 171,
    "templates": DEFAULT_TEMPLATES
}


def load_config():
    if os.path.exists(CONFIG_FILE):
        try:
            with open(CONFIG_FILE, "r", encoding="utf-8") as f:
                data = json.load(f)
                # Ensure templates are present
                if "templates" not in data or not data["templates"]:
                    data["templates"] = DEFAULT_TEMPLATES
                return data
        except Exception:
            pass
    return DEFAULT_CONFIG.copy()


def save_config(cfg):
    try:
        with open(CONFIG_FILE, "w", encoding="utf-8") as f:
            json.dump(cfg, f, indent=2)
        return True
    except Exception as e:
        print(f"Error saving config: {e}")
        return False


@app.route("/")
def index():
    return render_template("index.html")


@app.route("/api/config", methods=["GET", "POST"])
def config_api():
    if request.method == "GET":
        cfg = load_config()
        excel_status = None
        if cfg.get("master_excel_path") and os.path.exists(cfg["master_excel_path"]):
            excel_status = get_excel_info(cfg["master_excel_path"], cfg.get("selected_sheet"))
        
        estimate_status = None
        if cfg.get("estimate_excel_path") and os.path.exists(cfg["estimate_excel_path"]):
            estimate_status = inspect_estimate_excel(cfg["estimate_excel_path"])

        return jsonify({
            "config": cfg,
            "excel_status": excel_status,
            "estimate_status": estimate_status
        })

    elif request.method == "POST":
        data = request.json or {}
        cfg = load_config()
        for key in [
            "master_excel_path", "selected_sheet", "auto_increment_ref", 
            "default_survey_date_today", "create_backup",
            "estimate_excel_path", "estimate_selected_sheet", "estimate_start_row"
        ]:
            if key in data:
                cfg[key] = data[key]
        
        save_config(cfg)
        
        excel_status = None
        if cfg.get("master_excel_path") and os.path.exists(cfg["master_excel_path"]):
            excel_status = get_excel_info(cfg["master_excel_path"], cfg.get("selected_sheet"))

        estimate_status = None
        if cfg.get("estimate_excel_path") and os.path.exists(cfg["estimate_excel_path"]):
            estimate_status = inspect_estimate_excel(cfg["estimate_excel_path"])

        return jsonify({
            "success": True,
            "config": cfg,
            "excel_status": excel_status,
            "estimate_status": estimate_status
        })


@app.route("/api/list-local-excels", methods=["GET"])
def list_local_excels():
    """No-op endpoint: Full-PC recursive disk scanning has been removed to eliminate lag and crashes."""
    return jsonify({"success": True, "files": []})


def open_native_file_dialog(title="Select Excel Spreadsheet", initial_dir=None):
    """
    Opens real native Windows File Explorer dialog and returns the absolute file path selected by the user.
    Uses Win32 GetOpenFileNameW via ctypes (100% thread-safe on Flask worker threads),
    with PowerShell fallback. Tkinter is never invoked on worker threads to prevent native crashes.
    """
    if sys.platform == "win32":
        # 1. Thread-safe Win32 COM / Common Dialogs via ctypes
        try:
            import ctypes
            from ctypes import wintypes

            class OPENFILENAMEW(ctypes.Structure):
                _fields_ = [
                    ("lStructSize", wintypes.DWORD),
                    ("hwndOwner", wintypes.HWND),
                    ("hInstance", wintypes.HINSTANCE),
                    ("lpstrFilter", wintypes.LPCWSTR),
                    ("lpstrCustomFilter", wintypes.LPWSTR),
                    ("nMaxCustFilter", wintypes.DWORD),
                    ("nFilterIndex", wintypes.DWORD),
                    ("lpstrFile", wintypes.LPWSTR),
                    ("nMaxFile", wintypes.DWORD),
                    ("lpstrFileTitle", wintypes.LPWSTR),
                    ("nMaxFileTitle", wintypes.DWORD),
                    ("lpstrInitialDir", wintypes.LPCWSTR),
                    ("lpstrTitle", wintypes.LPCWSTR),
                    ("Flags", wintypes.DWORD),
                    ("nFileOffset", wintypes.WORD),
                    ("nFileExtension", wintypes.WORD),
                    ("lpstrDefExt", wintypes.LPCWSTR),
                    ("lCustData", wintypes.LPARAM),
                    ("lpfnHook", wintypes.LPVOID),
                    ("lpTemplateName", wintypes.LPCWSTR),
                    ("pvReserved", wintypes.LPVOID),
                    ("dwReserved", wintypes.DWORD),
                    ("FlagsEx", wintypes.DWORD)
                ]

            filter_str = "Excel Spreadsheets (*.xlsx; *.xls; *.xlsm)\0*.xlsx;*.xls;*.xlsm\0All Files (*.*)\0*.*\0\0"
            buffer = ctypes.create_unicode_buffer(1024)
            ofn = OPENFILENAMEW()
            ofn.lStructSize = ctypes.sizeof(OPENFILENAMEW)
            ofn.hwndOwner = None
            ofn.lpstrFilter = filter_str
            ofn.nFilterIndex = 1
            ofn.lpstrFile = ctypes.cast(buffer, wintypes.LPWSTR)
            ofn.nMaxFile = 1024
            ofn.lpstrInitialDir = initial_dir or os.path.expanduser("~/Desktop")
            ofn.lpstrTitle = title
            # OFN_FILEMUSTEXIST (0x1000) | OFN_PATHMUSTEXIST (0x800) | OFN_NOCHANGEDIR (0x8)
            ofn.Flags = 0x00001000 | 0x00000800 | 0x00000008

            if ctypes.windll.comdlg32.GetOpenFileNameW(ctypes.byref(ofn)):
                selected = buffer.value
                if selected and os.path.exists(selected):
                    return os.path.normpath(os.path.abspath(selected))
        except Exception:
            pass

        # 2. PowerShell STA OpenFileDialog fallback
        try:
            ps_script = (
                "[System.Reflection.Assembly]::LoadWithPartialName('System.windows.forms') | Out-Null; "
                "$f = New-Object System.Windows.Forms.OpenFileDialog; "
                f"$f.Title = '{title}'; "
                "$f.Filter = 'Excel Files (*.xlsx;*.xls;*.xlsm)|*.xlsx;*.xls;*.xlsm|All Files (*.*)|*.*'; "
                "$f.RestoreDirectory = $true; "
                "$t = New-Object System.Windows.Forms.Form; $t.TopMost = $true; "
                "if ($f.ShowDialog($t) -eq [System.Windows.Forms.DialogResult]::OK) { Write-Host $f.FileName }"
            )
            res = subprocess.run(
                ["powershell", "-NoProfile", "-STA", "-Command", ps_script],
                capture_output=True, text=True, timeout=60
            )
            out = res.stdout.strip()
            if out and os.path.exists(out):
                return os.path.normpath(os.path.abspath(out))
        except Exception:
            pass

    return ""


@app.route("/api/browse-native-excel", methods=["POST"])
def browse_native_excel():
    """Opens the real Windows file dialog and links directly to the user's actual spreadsheet on disk."""
    data = request.json or {}
    mode = data.get("type", "master")
    cfg = load_config()

    current_path = cfg.get("master_excel_path") if mode == "master" else cfg.get("estimate_excel_path")
    initial_dir = os.path.dirname(current_path) if current_path and os.path.exists(os.path.dirname(current_path)) else None

    title = "Select Master MIS Excel Spreadsheet" if mode == "master" else "Select Assessment / Estimate Excel Spreadsheet"
    file_path = open_native_file_dialog(title=title, initial_dir=initial_dir)

    if not file_path:
        return jsonify({"success": False, "cancelled": True, "error": "No file selected."})

    if not os.path.exists(file_path):
        return jsonify({"success": False, "error": f"Selected file does not exist: {file_path}"})

    ext = os.path.splitext(file_path)[1].lower()
    if ext not in [".xlsx", ".xls", ".xlsm"]:
        return jsonify({"success": False, "error": "Please select a valid Excel file (.xlsx, .xls, .xlsm)."})

    norm_path = os.path.normpath(os.path.abspath(file_path))
    mtime = os.path.getmtime(norm_path)
    size_kb = os.path.getsize(norm_path) / 1024

    if mode == "master":
        info = get_excel_info(norm_path)
        if not info.get("success"):
            return jsonify({"success": False, "error": info.get("error", "Failed to inspect Master Excel file.")})

        save_as_master = data.get("save_as_master", True)
        if save_as_master:
            cfg["master_excel_path"] = norm_path
            cfg["selected_sheet"] = info.get("selected_sheet", "")
            save_config(cfg)

        return jsonify({
            "success": True,
            "type": "master",
            "file_path": norm_path,
            "filename": os.path.basename(norm_path),
            "folder": os.path.dirname(norm_path),
            "last_modified": datetime.fromtimestamp(mtime).strftime("%d-%b-%Y %I:%M:%S %p"),
            "file_size": f"{size_kb:.1f} KB",
            "sheet": info.get("selected_sheet", ""),
            "selected_sheet": info.get("selected_sheet", ""),
            "sheets": info.get("sheets", []),
            "last_ref_no": info.get("last_ref_no"),
            "next_ref_no": info.get("next_ref_no"),
            "excel_status": info,
            "config": cfg
        })
    else:
        info = inspect_estimate_excel(norm_path)
        if not info.get("success"):
            return jsonify({"success": False, "error": info.get("error", "Failed to inspect Estimate Excel file.")})

        save_as_estimate = data.get("save_as_estimate", True)
        if save_as_estimate:
            cfg["estimate_excel_path"] = norm_path
            cfg["estimate_selected_sheet"] = info.get("selected_sheet", "")
            cfg["estimate_start_row"] = info.get("data_start_row", 171)
            save_config(cfg)

        return jsonify({
            "success": True,
            "type": "estimate",
            "file_path": norm_path,
            "filename": os.path.basename(norm_path),
            "folder": os.path.dirname(norm_path),
            "last_modified": datetime.fromtimestamp(mtime).strftime("%d-%b-%Y %I:%M:%S %p"),
            "file_size": f"{size_kb:.1f} KB",
            "sheet": info.get("selected_sheet", ""),
            "selected_sheet": info.get("selected_sheet", ""),
            "sheets": info.get("sheets", []),
            "data_start_row": info.get("data_start_row", 171),
            "excel_status": info,
            "config": cfg
        })


@app.route("/api/set-excel-path", methods=["POST"])
def set_excel_path():
    """Directly links an Excel spreadsheet path entered or pasted by the user."""
    data = request.json or {}
    raw_path = data.get("file_path", "").strip().strip('"').strip("'")
    mode = data.get("type", "master")
    sheet_name = data.get("sheet_name")

    if not raw_path:
        return jsonify({"success": False, "error": "No file path provided."})

    norm_path = os.path.normpath(os.path.abspath(raw_path))
    if not os.path.exists(norm_path):
        return jsonify({"success": False, "error": f"File does not exist at this path: {norm_path}"})

    ext = os.path.splitext(norm_path)[1].lower()
    if ext not in [".xlsx", ".xls", ".xlsm"]:
        return jsonify({"success": False, "error": "Please provide a valid Excel file (.xlsx, .xls, .xlsm)."})

    cfg = load_config()
    mtime = os.path.getmtime(norm_path)
    size_kb = os.path.getsize(norm_path) / 1024

    if mode == "master":
        info = get_excel_info(norm_path, target_sheet=sheet_name)
        if not info.get("success"):
            return jsonify({"success": False, "error": info.get("error", "Could not read Excel file.")})
        save_as_master = data.get("save_as_master", True)
        if save_as_master:
            cfg["master_excel_path"] = norm_path
            cfg["selected_sheet"] = info.get("selected_sheet", "")
            save_config(cfg)
        return jsonify({
            "success": True,
            "type": "master",
            "file_path": norm_path,
            "filename": os.path.basename(norm_path),
            "folder": os.path.dirname(norm_path),
            "last_modified": datetime.fromtimestamp(mtime).strftime("%d-%b-%Y %I:%M:%S %p"),
            "file_size": f"{size_kb:.1f} KB",
            "sheet": info.get("selected_sheet", ""),
            "selected_sheet": info.get("selected_sheet", ""),
            "sheets": info.get("sheets", []),
            "last_ref_no": info.get("last_ref_no"),
            "next_ref_no": info.get("next_ref_no"),
            "excel_status": info,
            "config": cfg
        })
    else:
        info = inspect_estimate_excel(norm_path, target_sheet=sheet_name)
        if not info.get("success"):
            return jsonify({"success": False, "error": info.get("error", "Could not inspect Estimate Excel file.")})
        save_as_estimate = data.get("save_as_estimate", True)
        if save_as_estimate:
            cfg["estimate_excel_path"] = norm_path
            cfg["estimate_selected_sheet"] = info.get("selected_sheet", "")
            cfg["estimate_start_row"] = info.get("data_start_row", 171)
            save_config(cfg)
        return jsonify({
            "success": True,
            "type": "estimate",
            "file_path": norm_path,
            "filename": os.path.basename(norm_path),
            "folder": os.path.dirname(norm_path),
            "last_modified": datetime.fromtimestamp(mtime).strftime("%d-%b-%Y %I:%M:%S %p"),
            "file_size": f"{size_kb:.1f} KB",
            "sheet": info.get("selected_sheet", ""),
            "selected_sheet": info.get("selected_sheet", ""),
            "sheets": info.get("sheets", []),
            "data_start_row": info.get("data_start_row", 171),
            "excel_status": info,
            "config": cfg
        })


@app.route("/api/excel-file-status", methods=["GET"])
def excel_file_status():
    """Returns disk status, modification timestamp, size, and latest ref for the active Excel file."""
    mode = request.args.get("type", "master")
    custom_path = request.args.get("path")
    cfg = load_config()

    file_path = custom_path or (cfg.get("master_excel_path") if mode == "master" else cfg.get("estimate_excel_path"))
    if not file_path or not os.path.exists(file_path):
        return jsonify({"success": False, "exists": False, "error": "No file linked or file does not exist on disk."})

    norm_path = os.path.normpath(os.path.abspath(file_path))
    mtime = os.path.getmtime(norm_path)
    size_kb = os.path.getsize(norm_path) / 1024

    if mode == "master":
        sheet = cfg.get("selected_sheet")
        info = get_excel_info(norm_path, target_sheet=sheet)
        return jsonify({
            "success": True,
            "exists": True,
            "type": "master",
            "file_path": norm_path,
            "filename": os.path.basename(norm_path),
            "folder": os.path.dirname(norm_path),
            "last_modified": datetime.fromtimestamp(mtime).strftime("%d-%b-%Y %I:%M:%S %p"),
            "file_size": f"{size_kb:.1f} KB",
            "selected_sheet": info.get("selected_sheet", ""),
            "sheets": info.get("sheets", []),
            "last_ref_no": info.get("last_ref_no", 0),
            "next_ref_no": info.get("next_ref_no", 1),
            "last_row": info.get("last_row", 1)
        })
    else:
        sheet = cfg.get("estimate_selected_sheet")
        info = inspect_estimate_excel(norm_path, target_sheet=sheet)
        return jsonify({
            "success": True,
            "exists": True,
            "type": "estimate",
            "file_path": norm_path,
            "filename": os.path.basename(norm_path),
            "folder": os.path.dirname(norm_path),
            "last_modified": datetime.fromtimestamp(mtime).strftime("%d-%b-%Y %I:%M:%S %p"),
            "file_size": f"{size_kb:.1f} KB",
            "selected_sheet": info.get("selected_sheet", ""),
            "sheets": info.get("sheets", []),
            "start_row": info.get("data_start_row", 171)
        })


@app.route("/api/show-in-folder", methods=["POST"])
def show_in_folder():
    """Opens Windows Explorer with the active Excel spreadsheet highlighted."""
    data = request.json or {}
    file_path = data.get("file_path")
    if not file_path:
        mode = data.get("type", "master")
        cfg = load_config()
        file_path = cfg.get("master_excel_path") if mode == "master" else cfg.get("estimate_excel_path")

    if not file_path or not os.path.exists(file_path):
        return jsonify({"success": False, "error": f"File does not exist: {file_path}"})

    try:
        norm_path = os.path.normpath(file_path)
        if sys.platform == "win32":
            subprocess.Popen(f'explorer /select,"{norm_path}"')
        else:
            subprocess.Popen(["xdg-open", os.path.dirname(norm_path)])
        return jsonify({"success": True})
    except Exception as e:
        return jsonify({"success": False, "error": str(e)})


@app.route("/api/upload-master-excel", methods=["POST"])
def upload_master_excel():
    """Receives an Excel file chosen via browser file dialog and sets it as the active Master Excel."""
    if "file" not in request.files:
        return jsonify({"success": False, "error": "No file uploaded."})
    
    file = request.files["file"]
    if not file or not file.filename:
        return jsonify({"success": False, "error": "No file selected."})

    ext = os.path.splitext(file.filename)[1].lower()
    if ext not in [".xlsx", ".xls", ".xlsm"]:
        return jsonify({"success": False, "error": "Please select a valid Excel file (.xlsx, .xls, .xlsm)."})

    # Check if a file with this exact filename exists in local project/Desktop/Downloads to link directly
    raw_name = file.filename
    clean_name = secure_filename(file.filename)
    candidates = [
        os.path.expanduser(f"~/Downloads/{raw_name}"),
        os.path.expanduser(f"~/Desktop/{raw_name}"),
        os.path.expanduser(f"~/Documents/{raw_name}"),
        os.path.abspath(os.path.join(BASE_DIR, raw_name)),
        os.path.abspath(os.path.join(BASE_DIR, "..", raw_name)),
        os.path.abspath(os.path.join(BASE_DIR, "test_data", raw_name))
    ]
    target_path = None
    for cand in candidates:
        if os.path.exists(cand):
            target_path = cand
            break

    is_copy = False
    if not target_path:
        master_dir = os.path.join(UPLOAD_FOLDER, "master_excel")
        os.makedirs(master_dir, exist_ok=True)
        target_path = os.path.abspath(os.path.join(master_dir, clean_name or "master.xlsx"))
        file.save(target_path)
        is_copy = True

    # Inspect the master excel
    info = get_excel_info(target_path)
    if not info.get("success"):
        return jsonify({"success": False, "error": info.get("error", "Could not read Excel file.")})

    cfg = load_config()
    cfg["master_excel_path"] = target_path
    cfg["selected_sheet"] = info.get("selected_sheet", "")
    save_config(cfg)

    mtime = os.path.getmtime(target_path)
    size_kb = os.path.getsize(target_path) / 1024

    return jsonify({
        "success": True,
        "file_path": target_path,
        "filename": os.path.basename(target_path),
        "folder": os.path.dirname(target_path),
        "last_modified": datetime.fromtimestamp(mtime).strftime("%d-%b-%Y %I:%M:%S %p"),
        "file_size": f"{size_kb:.1f} KB",
        "is_copy": is_copy,
        "config": cfg,
        "excel_status": info
    })


@app.route("/api/rescan-ref", methods=["GET"])
def rescan_ref():
    """Re-scans the configured Master Excel to fetch latest & next Ref No."""
    cfg = load_config()
    master_path = cfg.get("master_excel_path")
    selected_sheet = cfg.get("selected_sheet")

    if not master_path or not os.path.exists(master_path):
        return jsonify({"success": False, "error": "Master Excel file is not configured or does not exist."})

    info = get_excel_info(master_path, selected_sheet)
    if not info.get("success"):
        return jsonify({"success": False, "error": info.get("error", "Failed to scan Master Excel.")})

    return jsonify({
        "success": True,
        "last_ref_no": info.get("last_ref_no", 0),
        "next_ref_no": info.get("next_ref_no", 1),
        "last_row": info.get("last_row", 1),
        "selected_sheet": info.get("selected_sheet")
    })


@app.route("/api/detect-excel", methods=["POST"])
def detect_excel():
    data = request.json or {}
    path = data.get("file_path", "").strip().strip('"')
    sheet_name = data.get("sheet_name", "").strip()

    if not path:
        return jsonify({"success": False, "error": "Please provide an Excel file path."})
    if not os.path.exists(path):
        return jsonify({"success": False, "error": f"File does not exist: {path}"})

    info = get_excel_info(path, sheet_name if sheet_name else None)
    return jsonify(info)


@app.route("/api/extract", methods=["POST"])
def extract_pdfs():
    if "files" not in request.files:
        return jsonify({"success": False, "error": "No files uploaded."})

    uploaded_files = request.files.getlist("files")
    if not uploaded_files or uploaded_files[0].filename == "":
        return jsonify({"success": False, "error": "No files selected."})

    cfg = load_config()
    master_path = cfg.get("master_excel_path")
    selected_sheet = cfg.get("selected_sheet")

    # Determine starting Ref No from master excel if available
    next_ref = 1
    if master_path and os.path.exists(master_path):
        info = get_excel_info(master_path, selected_sheet)
        if info.get("success"):
            next_ref = info.get("next_ref_no", 1)

    extracted_records = []
    errors = []

    for file in uploaded_files:
        if not file.filename.lower().endswith(".pdf"):
            continue

        filename = secure_filename(file.filename)
        save_path = os.path.join(UPLOAD_FOLDER, filename)
        file.save(save_path)

        try:
            record = extract_data_from_pdf(save_path, cfg.get("templates"))
            record["ref_no"] = next_ref
            next_ref += 1
            extracted_records.append(record)
        except Exception as e:
            errors.append({"filename": file.filename, "error": str(e)})

    return jsonify({
        "success": True,
        "records": extracted_records,
        "errors": errors,
        "total_processed": len(extracted_records)
    })


@app.route("/api/append", methods=["POST"])
def append_to_excel():
    data = request.json or {}
    rows_data = data.get("rows", [])

    if not rows_data:
        return jsonify({"success": False, "error": "No records provided to append."})

    cfg = load_config()
    file_path = data.get("file_path") or cfg.get("master_excel_path")
    sheet_name = data.get("sheet_name") or cfg.get("selected_sheet")
    create_backup = cfg.get("create_backup", True)

    if not file_path or not os.path.exists(file_path):
        return jsonify({
            "success": False,
            "error": f"Master Excel file does not exist on disk: {file_path}. Please browse or set the correct path."
        })

    result = append_rows_to_excel(file_path, sheet_name, rows_data, create_backup=create_backup)
    if result.get("success"):
        norm_path = os.path.normpath(os.path.abspath(file_path))
        mtime = os.path.getmtime(norm_path)
        size_kb = os.path.getsize(norm_path) / 1024
        result["file_path"] = norm_path
        result["filename"] = os.path.basename(norm_path)
        result["folder"] = os.path.dirname(norm_path)
        result["last_modified"] = datetime.fromtimestamp(mtime).strftime("%d-%b-%Y %I:%M:%S %p")
        result["file_size"] = f"{size_kb:.1f} KB"
        result["sheet"] = sheet_name or "Active Sheet"
    return jsonify(result)


@app.route("/api/open-excel", methods=["POST"])
def open_excel():
    data = request.json or {}
    cfg = load_config()
    file_path = data.get("file_path") or cfg.get("master_excel_path")

    if not file_path or not os.path.exists(file_path):
        return jsonify({"success": False, "error": f"Master Excel file not found: {file_path}"})

    try:
        norm_path = os.path.normpath(file_path)
        if sys.platform == "win32":
            os.startfile(norm_path)
        else:
            subprocess.Popen(["xdg-open", norm_path])
        return jsonify({"success": True, "file_path": norm_path})
    except Exception as e:
        return jsonify({"success": False, "error": f"Failed to open Excel: {str(e)}"})


@app.route("/api/templates", methods=["GET", "POST"])
def templates_api():
    cfg = load_config()
    if request.method == "GET":
        return jsonify({"templates": cfg.get("templates", DEFAULT_TEMPLATES)})

    elif request.method == "POST":
        new_template = request.json
        if not new_template or not new_template.get("name"):
            return jsonify({"success": False, "error": "Template must have a name."})

        templates = cfg.get("templates", [])
        # Update if exists, else append
        existing = False
        for i, t in enumerate(templates):
            if t.get("id") == new_template.get("id"):
                templates[i] = new_template
                existing = True
                break

        if not existing:
            templates.append(new_template)

        cfg["templates"] = templates
        save_config(cfg)
        return jsonify({"success": True, "templates": templates})


# ============================================================
#  ESTIMATE PARTS UPDATER ENDPOINTS
# ============================================================

@app.route("/api/extract-estimate", methods=["POST"])
def extract_estimate():
    """Extracts line items (Parts and Labour) and metadata from an uploaded bill estimate PDF."""
    if "file" not in request.files and "files" not in request.files:
        return jsonify({"success": False, "error": "No file uploaded."})
    
    file = request.files.get("file") or (request.files.getlist("files")[0] if request.files.getlist("files") else None)
    if not file or not file.filename:
        return jsonify({"success": False, "error": "No file selected."})

    if not file.filename.lower().endswith(".pdf"):
        return jsonify({"success": False, "error": "Please select a valid PDF file."})

    filename = secure_filename(file.filename)
    estimate_uploads = os.path.join(UPLOAD_FOLDER, "estimates")
    os.makedirs(estimate_uploads, exist_ok=True)
    save_path = os.path.join(estimate_uploads, filename)
    file.save(save_path)

    try:
        result = parse_estimate_pdf(save_path)
        return jsonify(result)
    except Exception as e:
        return jsonify({"success": False, "error": f"Failed to parse bill estimate: {str(e)}"})


@app.route("/api/inspect-estimate-excel", methods=["POST"])
def inspect_estimate_excel_api():
    """Inspects target Excel spreadsheet for PARTS COMPUTATION table location."""
    data = request.json or {}
    file_path = data.get("file_path", "").strip().strip('"')
    if not file_path:
        return jsonify({"success": False, "error": "Please provide an Excel file path."})
    if not os.path.exists(file_path):
        return jsonify({"success": False, "error": f"File not found: {file_path}"})

    info = inspect_estimate_excel(file_path)
    return jsonify(info)


@app.route("/api/upload-estimate-excel", methods=["POST"])
def upload_estimate_excel():
    """Allows browsing / uploading an Estimate Master Format Excel file."""
    if "file" not in request.files:
        return jsonify({"success": False, "error": "No file uploaded."})
    
    file = request.files["file"]
    if not file or not file.filename:
        return jsonify({"success": False, "error": "No file selected."})

    ext = os.path.splitext(file.filename)[1].lower()
    if ext not in [".xlsx", ".xls", ".xlsm"]:
        return jsonify({"success": False, "error": "Please select a valid Excel file (.xlsx, .xls, .xlsm)."})

    raw_name = file.filename
    clean_name = secure_filename(file.filename)
    candidates = [
        os.path.expanduser(f"~/Downloads/{raw_name}"),
        os.path.expanduser(f"~/Desktop/{raw_name}"),
        os.path.expanduser(f"~/Documents/{raw_name}"),
        os.path.abspath(os.path.join(os.path.dirname(__file__), raw_name)),
        os.path.abspath(os.path.join(os.path.dirname(__file__), "..", raw_name)),
        os.path.abspath(os.path.join(os.path.dirname(__file__), "test_data", raw_name))
    ]
    target_path = None
    for cand in candidates:
        if os.path.exists(cand):
            target_path = cand
            break

    if not target_path:
        master_dir = os.path.join(UPLOAD_FOLDER, "estimate_excel")
        os.makedirs(master_dir, exist_ok=True)
        target_path = os.path.abspath(os.path.join(master_dir, clean_name or "estimate_format.xls"))
        file.save(target_path)

    info = inspect_estimate_excel(target_path)
    if not info.get("success"):
        return jsonify({"success": False, "error": info.get("error", "Could not read Excel file.")})

    cfg = load_config()
    cfg["estimate_excel_path"] = target_path
    cfg["estimate_selected_sheet"] = info.get("selected_sheet", "")
    cfg["estimate_start_row"] = info.get("data_start_row", 171)
    save_config(cfg)

    return jsonify({
        "success": True,
        "file_path": target_path,
        "filename": os.path.basename(target_path),
        "config": cfg,
        "excel_status": info
    })


@app.route("/api/write-estimate", methods=["POST"])
def write_estimate_api():
    """Writes selected parts descriptions (Col C) and amounts (Col D) into the target sheet."""
    try:
        data = request.json or {}
        cfg = load_config()
        
        file_path = data.get("file_path") or cfg.get("estimate_excel_path") or cfg.get("master_excel_path")
        sheet_name = data.get("sheet_name") or cfg.get("estimate_selected_sheet")
        parts = data.get("parts", [])
        start_row = int(data.get("start_row") or cfg.get("estimate_start_row") or 171)
        bill_no = data.get("bill_no", "")

        if not file_path or not os.path.exists(file_path):
            return jsonify({"success": False, "error": f"Target Estimate Excel file not specified or does not exist: {file_path}"})

        res = write_parts_to_excel(
            file_path=file_path,
            sheet_name=sheet_name,
            parts=parts,
            start_row=start_row,
            bill_no=bill_no
        )

        if res.get("success"):
            norm_path = os.path.normpath(os.path.abspath(file_path))
            mtime = os.path.getmtime(norm_path)
            size_kb = os.path.getsize(norm_path) / 1024
            res["file_path"] = norm_path
            res["filename"] = os.path.basename(norm_path)
            res["folder"] = os.path.dirname(norm_path)
            res["last_modified"] = datetime.fromtimestamp(mtime).strftime("%d-%b-%Y %I:%M:%S %p")
            res["file_size"] = f"{size_kb:.1f} KB"
            # Auto-update next start row and last written file in config
            cfg["estimate_excel_path"] = norm_path
            if res.get("sheet"):
                cfg["estimate_selected_sheet"] = res["sheet"]
            if res.get("end_row"):
                cfg["estimate_start_row"] = res["end_row"] + 1
            save_config(cfg)

        return jsonify(res)
    except Exception as e:
        return jsonify({"success": False, "error": f"Error writing parts to Excel: {str(e)}"})


@app.route("/api/open-estimate-excel", methods=["POST"])
def open_estimate_excel():
    """Opens the active estimate excel spreadsheet in Microsoft Excel."""
    data = request.json or {}
    cfg = load_config()
    file_path = data.get("file_path") or cfg.get("estimate_excel_path") or cfg.get("master_excel_path")

    if not file_path or not os.path.exists(file_path):
        return jsonify({"success": False, "error": f"Estimate Excel file not found: {file_path}"})

    try:
        if sys.platform == "win32":
            os.startfile(file_path)
        else:
            subprocess.Popen(["xdg-open", file_path])
        return jsonify({"success": True, "file_path": file_path})
    except Exception as e:
        return jsonify({"success": False, "error": f"Failed to open Excel: {str(e)}"})


# ========================================================
#  WEB CLAIM PORTAL & EXTENSION ENDPOINTS
# ========================================================

@app.route("/test-portal")
def test_portal():
    """Renders simulated Universal Sompo claim portal for testing Chrome extension."""
    return render_template("test_portal.html")


@app.route("/api/web-claim", methods=["POST", "OPTIONS"])
def receive_web_claim():
    """Receives claim record from Chrome extension or web parser and queues it."""
    if request.method == "OPTIONS":
        return jsonify({"success": True})

    data = request.json or {}
    record = data.get("record") or data
    if not record:
        return jsonify({"success": False, "error": "No claim record provided."})

    cfg = load_config()
    master_path = cfg.get("master_excel_path")
    selected_sheet = cfg.get("selected_sheet")

    # Determine starting Ref No from master excel if available
    next_ref = 1
    if master_path and os.path.exists(master_path):
        info = get_excel_info(master_path, selected_sheet)
        if info.get("success"):
            next_ref = info.get("next_ref_no", 1)

    with web_claims_lock:
        if web_claims_queue:
            last_ref = max([r.get("ref_no", 0) for r in web_claims_queue] + [next_ref - 1])
            record["ref_no"] = last_ref + 1
        else:
            record["ref_no"] = next_ref

        record["received_at"] = datetime.now().strftime("%I:%M:%S %p")
        web_claims_queue.append(record)
        total_q = len(web_claims_queue)

    return jsonify({
        "success": True,
        "ref_no": record["ref_no"],
        "total_queued": total_q,
        "record": record
    })


@app.route("/api/web-claims", methods=["GET"])
def get_web_claims():
    """Returns all queued web claims."""
    with web_claims_lock:
        records_copy = list(web_claims_queue)
    return jsonify({
        "success": True,
        "records": records_copy,
        "total": len(records_copy)
    })


@app.route("/api/clear-web-claims", methods=["POST"])
def clear_web_claims():
    """Clears the web claims queue."""
    with web_claims_lock:
        web_claims_queue.clear()
    return jsonify({"success": True})


if __name__ == "__main__":
    print("\n" + "=" * 60)
    print("  INSURANCE CLAIM MIS EXTRACTOR - LOCALHOST SERVER")
    print("  Open in browser: http://127.0.0.1:5055")
    print("=" * 60 + "\n")
    app.run(host="127.0.0.1", port=5055, debug=False)
