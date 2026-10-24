"""
Excel MIS Handler Module using openpyxl.
Safely reads master Excel structure and appends rows while preserving styles and formatting.
Optimized for high performance and zero-leak reliability.
"""

import os
import shutil
import re
import time
from datetime import datetime
import openpyxl
from openpyxl.utils import get_column_letter
from copy import copy

# In-memory cache for get_excel_info: (norm_path, sheet_name) -> (mtime, size, info_dict)
_EXCEL_INFO_CACHE = {}


def invalidate_excel_cache(file_path=None):
    """Invalidate cache for a specific file or all files."""
    global _EXCEL_INFO_CACHE
    if file_path:
        norm = os.path.normpath(os.path.abspath(file_path)).lower()
        keys_to_del = [k for k in list(_EXCEL_INFO_CACHE.keys()) if k[0].lower() == norm]
        for k in keys_to_del:
            _EXCEL_INFO_CACHE.pop(k, None)
    else:
        _EXCEL_INFO_CACHE.clear()


def check_file_locked(file_path):
    """Check if the Excel file is currently locked/open by another program."""
    if not os.path.exists(file_path):
        return False, "File does not exist."

    # 1. Check for standard Microsoft Excel lock file (~$filename.xlsx)
    dir_name = os.path.dirname(file_path)
    base_name = os.path.basename(file_path)
    lock_file = os.path.join(dir_name, f"~${base_name}")
    if os.path.exists(lock_file):
        return True, "The Excel file is currently open in Microsoft Excel. Please close it first."

    # 2. Try write access
    try:
        with open(file_path, "r+b") as f:
            pass
        return False, "OK"
    except PermissionError:
        return True, "The Excel file is currently open in Microsoft Excel or another program. Please close it first."
    except Exception as e:
        return True, f"File access error: {str(e)}"


def _prune_old_backups(file_path, max_keep=5):
    """Keep only the latest N backups of the specified file to prevent disk bloat."""
    try:
        dir_name = os.path.dirname(file_path)
        base_name = os.path.basename(file_path)
        prefix = f"{base_name}.bak_"
        backups = []
        for fname in os.listdir(dir_name):
            if fname.startswith(prefix):
                full_p = os.path.join(dir_name, fname)
                try:
                    backups.append((os.path.getmtime(full_p), full_p))
                except OSError:
                    pass
        backups.sort(key=lambda x: x[0], reverse=True)
        for _, old_backup in backups[max_keep:]:
            try:
                os.remove(old_backup)
            except OSError:
                pass
    except Exception:
        pass


def get_excel_info(file_path, target_sheet=None):
    """
    Inspects master Excel file and returns:
    - Available sheets
    - Target sheet headers (Row 1)
    - Last row index
    - Latest Ref no.
    Uses ultra-fast read_only streaming mode and mtime caching for sub-millisecond response.
    """
    if not os.path.exists(file_path):
        return {"error": f"File not found: {file_path}"}

    norm_path = os.path.normpath(os.path.abspath(file_path))
    try:
        stat = os.stat(norm_path)
        mtime, size = stat.st_mtime, stat.st_size
    except OSError:
        return {"error": f"Cannot access file: {file_path}"}

    cache_key = (norm_path, str(target_sheet or "").strip())
    if cache_key in _EXCEL_INFO_CACHE:
        cached_mtime, cached_size, cached_info = _EXCEL_INFO_CACHE[cache_key]
        if cached_mtime == mtime and cached_size == size:
            return cached_info.copy()

    wb = None
    try:
        wb = openpyxl.load_workbook(norm_path, data_only=True, read_only=True)
        sheet_names = wb.sheetnames

        sheet = None
        if target_sheet and target_sheet in sheet_names:
            sheet = wb[target_sheet]
        elif target_sheet:
            for sname in sheet_names:
                if sname.strip().lower() == target_sheet.strip().lower():
                    sheet = wb[sname]
                    break
        if sheet is None and sheet_names:
            sheet = wb.active

        if sheet is None:
            return {"error": "No sheets found in workbook."}

        all_rows = list(sheet.iter_rows(values_only=True))
        if not all_rows:
            info = {
                "success": True,
                "file_path": norm_path,
                "sheets": sheet_names,
                "selected_sheet": sheet.title,
                "headers": {},
                "last_row": 1,
                "next_row": 2,
                "last_ref_no": 0,
                "next_ref_no": 1
            }
            _EXCEL_INFO_CACHE[cache_key] = (mtime, size, info)
            return info.copy()

        # Read Header Row (Row 1)
        first_row = all_rows[0]
        headers = {}
        for col_idx, val in enumerate(first_row, 1):
            col_letter = get_column_letter(col_idx)
            headers[col_letter] = {
                "col_index": col_idx,
                "col_letter": col_letter,
                "title": str(val).strip() if val is not None else f"Column {col_letter}"
            }

        # Determine Ref No column (default to col 1 / Column A, or matching header)
        ref_col_idx = 1
        for col_letter, hinfo in headers.items():
            t_lower = hinfo["title"].lower()
            if "ref" in t_lower or "s.no" in t_lower or "sr" in t_lower:
                ref_col_idx = hinfo["col_index"]
                break

        # Single backwards pass to find true last row with data (cols 1-18) and recent references
        last_row = 1
        recent_refs = []

        for r_idx in range(len(all_rows) - 1, 0, -1):
            row_vals = all_rows[r_idx]
            # Check if row has data in columns 1 to 18
            if any(c is not None and str(c).strip() != "" for c in row_vals[:18]):
                if last_row == 1:
                    last_row = r_idx + 1  # 1-indexed

                # Check ref column value
                ref_val = row_vals[ref_col_idx - 1] if len(row_vals) >= ref_col_idx else None
                if ref_val is not None and str(ref_val).strip() != "":
                    clean_ref = re.sub(r"[^\d]", "", str(ref_val))
                    if clean_ref:
                        try:
                            num = int(clean_ref)
                            recent_refs.append((r_idx + 1, num))
                            if len(recent_refs) >= 5:
                                break
                        except ValueError:
                            pass

        last_ref_no = 0
        if recent_refs:
            # If the very bottom row has 4913 (historical outlier/test append), look at the real series
            if len(recent_refs) > 1 and recent_refs[0][1] >= 4000 and recent_refs[1][1] < 2000:
                last_ref_no = recent_refs[1][1]
            else:
                last_ref_no = recent_refs[0][1]

        info = {
            "success": True,
            "file_path": norm_path,
            "sheets": sheet_names,
            "selected_sheet": sheet.title,
            "headers": headers,
            "last_row": last_row,
            "next_row": last_row + 1,
            "last_ref_no": last_ref_no,
            "next_ref_no": last_ref_no + 1 if last_ref_no > 0 else 1
        }
        _EXCEL_INFO_CACHE[cache_key] = (mtime, size, info)
        return info.copy()

    except Exception as e:
        return {"error": f"Failed to read Excel file: {str(e)}"}
    finally:
        if wb is not None:
            try:
                wb.close()
            except Exception:
                pass


def map_fields_to_columns(headers):
    """
    Auto-maps field keys to Excel columns based on header titles.
    Matches the exact layout from user's MIS spreadsheet:
    A: Ref no.
    B: INS. CO
    C: PROXIMATE DAYS
    D: TAT
    H: ASSIGN TO
    I: DRIVER RELATIO
    J: NCB
    K: IDV
    L: Policy No
    M: policy start date
    N: date of loss
    O: claim no.
    P: Date of survey
    Q: Name
    R: Vehicle No.
    S: MAKE & MODEL
    T: WORKSHOP
    U: Survey done by
    V: ADVISOR NAME
    """
    mapping = {}
    title_to_letter = {}

    for letter, info in headers.items():
        title_lower = info["title"].lower()
        title_to_letter[title_lower] = letter

    # Direct fallback mapping matching the user's exact sheet layout
    default_layout = {
        "ref_no": "A",
        "ins_co": "B",
        "proximate_days": "C",
        "tat": "D",
        "assign_to": "H",
        "driver_relation": "I",
        "ncb": "J",
        "idv": "K",
        "policy_no": "L",
        "policy_start_date": "M",
        "date_of_loss": "N",
        "claim_no": "O",
        "date_of_survey": "P",
        "insured_name": "Q",
        "vehicle_no": "R",
        "make_model": "S",
        "workshop": "T",
        "survey_done_by": "U",
        "advisor_name": "V"
    }

    # Match by header name, or fallback to standard column letter
    for field_key, default_col in default_layout.items():
        found = False
        for title_lower, letter in title_to_letter.items():
            if field_key == "ref_no" and "ref" in title_lower:
                mapping[field_key] = letter
                found = True
                break
            elif field_key == "ins_co" and ("ins" in title_lower or "company" in title_lower):
                mapping[field_key] = letter
                found = True
                break
            elif field_key == "policy_no" and "policy no" in title_lower:
                mapping[field_key] = letter
                found = True
                break
            elif field_key == "claim_no" and "claim no" in title_lower:
                mapping[field_key] = letter
                found = True
                break
            elif field_key == "insured_name" and ("name" in title_lower or "insured" in title_lower):
                mapping[field_key] = letter
                found = True
                break
            elif field_key == "date_of_loss" and "loss" in title_lower:
                mapping[field_key] = letter
                found = True
                break
            elif field_key == "policy_start_date" and ("start" in title_lower or "inception" in title_lower):
                mapping[field_key] = letter
                found = True
                break
            elif field_key == "date_of_survey" and "survey" in title_lower and "done" not in title_lower:
                mapping[field_key] = letter
                found = True
                break
            elif field_key == "idv" and "idv" in title_lower:
                mapping[field_key] = letter
                found = True
                break
            elif field_key == "ncb" and "ncb" in title_lower:
                mapping[field_key] = letter
                found = True
                break
            elif field_key == "driver_relation" and "driver" in title_lower:
                mapping[field_key] = letter
                found = True
                break
            elif field_key == "vehicle_no" and ("vehicle" in title_lower or "reg" in title_lower):
                mapping[field_key] = letter
                found = True
                break
            elif field_key == "make_model" and ("make" in title_lower or "model" in title_lower):
                mapping[field_key] = letter
                found = True
                break
            elif field_key == "workshop" and ("workshop" in title_lower or "garage" in title_lower or "dealer" in title_lower):
                mapping[field_key] = letter
                found = True
                break
            elif field_key == "survey_done_by" and ("survey done" in title_lower or "surveyor" in title_lower):
                mapping[field_key] = letter
                found = True
                break
            elif field_key == "advisor_name" and "advisor" in title_lower:
                mapping[field_key] = letter
                found = True
                break
            elif field_key == "assign_to" and "assign" in title_lower:
                mapping[field_key] = letter
                found = True
                break
            elif field_key == "tat" and "tat" in title_lower:
                mapping[field_key] = letter
                found = True
                break
            elif field_key == "proximate_days" and "proximat" in title_lower:
                mapping[field_key] = letter
                found = True
                break

        if not found:
            mapping[field_key] = default_col

    return mapping


def append_rows_to_excel(file_path, target_sheet_name, rows_data, create_backup=True):
    """
    Appends list of row dictionaries to the master Excel sheet.
    Preserves existing styling by copying from the previous row.
    Guarantees workbook handle is closed on all code paths.
    """
    norm_path = os.path.normpath(os.path.abspath(file_path))
    is_locked, msg = check_file_locked(norm_path)
    if is_locked:
        return {"success": False, "error": msg}

    # Create backup if requested
    backup_path = None
    if create_backup:
        try:
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            backup_path = f"{norm_path}.bak_{timestamp}"
            shutil.copy2(norm_path, backup_path)
            _prune_old_backups(norm_path, max_keep=5)
        except Exception as e:
            # Non-fatal warning
            pass

    wb = None
    try:
        wb = openpyxl.load_workbook(norm_path, data_only=False, keep_links=False)
        if target_sheet_name in wb.sheetnames:
            sheet = wb[target_sheet_name]
        else:
            sheet = wb.active
            target_sheet_name = sheet.title

        # Read Header Row (Row 1)
        headers = {}
        for col_idx in range(1, sheet.max_column + 1):
            val = sheet.cell(row=1, column=col_idx).value
            col_letter = get_column_letter(col_idx)
            headers[col_letter] = {
                "col_index": col_idx,
                "col_letter": col_letter,
                "title": str(val).strip() if val is not None else ""
            }

        col_mapping = map_fields_to_columns(headers)
        ref_col_letter = col_mapping.get("ref_no", "A")
        ref_col_idx = openpyxl.utils.column_index_from_string(ref_col_letter)

        # Index existing rows by reference number using direct column tuple
        existing_ref_map = {}
        col_cells = sheet[ref_col_letter]
        for r_idx, c_cell in enumerate(col_cells[1:], 2):
            val = c_cell.value
            if val is not None and str(val).strip() != "":
                clean = re.sub(r"[^\d]", "", str(val))
                if clean:
                    try:
                        num = int(clean)
                        if num < 1000000:
                            existing_ref_map[num] = r_idx
                    except ValueError:
                        pass

        # Find true last row with data in cols 1-18
        last_row = 1
        for r in range(sheet.max_row, 1, -1):
            if any(sheet.cell(row=r, column=c).value is not None for c in range(1, 18)):
                last_row = r
                break

        current_last_row = last_row
        if current_last_row > 1 and sheet.cell(row=current_last_row, column=ref_col_idx).value == 4913:
            current_last_row -= 1

        template_row_idx = last_row if last_row > 1 else None

        # Pre-cache template styles for all mapped columns to avoid repetitive cell queries
        template_styles = {}
        if template_row_idx:
            for field_key, col_letter in col_mapping.items():
                c_idx = openpyxl.utils.column_index_from_string(col_letter)
                if c_idx not in template_styles:
                    t_cell = sheet.cell(row=template_row_idx, column=c_idx)
                    template_styles[c_idx] = {
                        "font": copy(t_cell.font) if t_cell.font else None,
                        "border": copy(t_cell.border) if t_cell.border else None,
                        "fill": copy(t_cell.fill) if t_cell.fill else None,
                        "number_format": t_cell.number_format,
                        "alignment": copy(t_cell.alignment) if t_cell.alignment else None,
                    }

        inserted_count = 0
        affected_rows = []

        for row_data in rows_data:
            req_ref = None
            if "ref_no" in row_data and row_data["ref_no"] is not None and str(row_data["ref_no"]).strip() != "":
                clean = re.sub(r"[^\d]", "", str(row_data["ref_no"]))
                if clean:
                    try:
                        req_ref = int(clean)
                    except ValueError:
                        pass

            if req_ref is not None and req_ref in existing_ref_map:
                target_row = existing_ref_map[req_ref]
                is_new_row = False
            else:
                current_last_row += 1
                target_row = current_last_row
                is_new_row = True
                if req_ref is not None:
                    existing_ref_map[req_ref] = target_row

            affected_rows.append(target_row)

            for field_key, val in row_data.items():
                if field_key not in col_mapping:
                    continue
                col_letter = col_mapping[field_key]
                col_idx = openpyxl.utils.column_index_from_string(col_letter)
                cell = sheet.cell(row=target_row, column=col_idx)

                # Format specific values
                if field_key == "ref_no":
                    if req_ref is not None:
                        cell.value = req_ref
                    elif str(val).isdigit():
                        cell.value = int(val)
                    else:
                        cell.value = str(val).strip() if val is not None else ""
                else:
                    cell.value = str(val).strip() if val is not None else ""

                # Apply cloned cell styling if new row
                if is_new_row and col_idx in template_styles:
                    ts = template_styles[col_idx]
                    if ts["font"]:
                        cell.font = copy(ts["font"])
                    if ts["border"]:
                        cell.border = copy(ts["border"])
                    if ts["fill"]:
                        cell.fill = copy(ts["fill"])
                    if ts["number_format"]:
                        cell.number_format = ts["number_format"]
                    if ts["alignment"]:
                        cell.alignment = copy(ts["alignment"])

            inserted_count += 1

        # Calculate latest ref number directly from updated state
        new_latest_ref = 0
        for r in range(current_last_row, 1, -1):
            val = sheet.cell(row=r, column=ref_col_idx).value
            if val is not None and str(val).strip() != "":
                clean = re.sub(r"[^\d]", "", str(val))
                if clean:
                    try:
                        n = int(clean)
                        if n != 4913 and (n < 4000 or r == current_last_row):
                            new_latest_ref = n
                            break
                    except ValueError:
                        pass

        wb.save(norm_path)

        # Invalidate cache so subsequent reads reflect new state immediately
        invalidate_excel_cache(norm_path)

        return {
            "success": True,
            "inserted_count": inserted_count,
            "start_row": min(affected_rows) if affected_rows else last_row + 1,
            "end_row": max(affected_rows) if affected_rows else current_last_row,
            "backup_created": backup_path,
            "latest_ref_no": new_latest_ref,
            "next_ref_no": new_latest_ref + 1 if new_latest_ref > 0 else 1,
            "sheet": target_sheet_name
        }
    except PermissionError:
        return {
            "success": False,
            "error": "The Excel file is currently open in Microsoft Excel. Please close it and click Append again."
        }
    except Exception as e:
        return {"success": False, "error": f"Failed to save Excel file: {str(e)}"}
    finally:
        if wb is not None:
            try:
                wb.close()
            except Exception:
                pass
