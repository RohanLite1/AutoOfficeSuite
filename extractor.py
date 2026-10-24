"""
PDF Data Extraction Engine for Insurance Claim Intimations & Policies.
Extracts data from digital PDFs using pdfplumber and regex/anchor patterns.
"""

import re
import os
import pdfplumber
from datetime import datetime

# Built-in Default Templates for Common Insurers
DEFAULT_TEMPLATES = [
    {
        "id": "bajaj",
        "name": "Bajaj General / Bajaj Allianz",
        "ins_co_code": "BAJAJ",
        "keywords": ["bajaj general insurance", "bajaj allianz", "bajaj caringly yours"],
        "fields": {
            "policy_no": {
                "anchors": ["Policy Reference Number", "Policy Number", "Policy No"],
                "regex": r"OG[-\d]+"
            },
            "claim_no": {
                "anchors": ["Claim Reference Number", "Claim Number", "Claim No"],
                "regex": r"[A-Z0-9]{2}[-\d]+"
            },
            "insured_name": {
                "anchors": ["Policy Holder\\Insured Name", "Policy Holder / Insured Name", "Policy Holder", "Insured Name"],
            },
            "date_of_loss": {
                "anchors": ["Date of Loss", "Loss Date"],
                "is_date": True
            },
            "policy_start_date": {
                "anchors": ["Policy Inception Date", "Policy Start Date", "Period of Insurance From"],
                "is_date": True
            },
            "ncb": {
                "anchors": ["No Claim Bonus", "NCB"],
                "regex": r"\b(\d{1,2})\b"
            },
            "driver_relation": {
                "anchors": ["Driver Relation with Insured", "Driver Relation"],
            },
            "vehicle_no": {
                "anchors": ["Vehicle Registration Number", "Registration Number", "Reg No", "Vehicle No"],
            },
            "vehicle_make": {
                "anchors": ["Vehicle Make", "Make"],
            },
            "vehicle_model": {
                "anchors": ["Vehicle Model", "Model"],
            },
            "make_model": {
                "anchors": ["Vehicle Make & Model", "Make & Model", "Make/Model"],
            },
            "workshop": {
                "anchors": ["Inspection Address", "Workshop Name", "Garage Name", "Repairer Name", "Dealer Name"],
            },
            "survey_done_by": {
                "anchors": ["Survey Done By", "Surveyor Name", "Deputed Surveyor", "Surveyor"],
            },
            "advisor_name": {
                "anchors": ["Advisor Name", "Service Advisor", "Claim Advisor"],
            },
            "assign_to": {
                "anchors": ["Assign To", "Assigned To", "Aassign To"],
            },
            "tat": {
                "anchors": ["Turn Around Time", "TAT Days"],
                "regex": r"\b(\d+)\b"
            },
            "proximate_days": {
                "anchors": ["Proximate Days"],
            },
            "idv": {
                "anchors": ["IDV", "Insured Declared Value"],
                "regex": r"([0-9,]+(?:/-)?)"
            }
        }
    },
    {
        "id": "nia",
        "name": "New India Assurance (NIA)",
        "ins_co_code": "NIA",
        "keywords": ["new india assurance", "the new india", "nia"],
        "fields": {
            "policy_no": {
                "anchors": ["Policy Number", "Policy No.", "Policy No"],
                "regex": r"([A-Z0-9\/-]{10,})"
            },
            "claim_no": {
                "anchors": ["Claim Number", "Claim No.", "Claim No"],
                "regex": r"([A-Z0-9\/-]{10,})"
            },
            "insured_name": {
                "anchors": ["Name of Insured", "Insured Name", "Policyholder Name"],
            },
            "date_of_loss": {
                "anchors": ["Date of Accident", "Date of Loss", "Time & Date of Loss"],
                "is_date": True
            },
            "policy_start_date": {
                "anchors": ["Policy Inception Date", "Period of Insurance From", "Policy From"],
                "is_date": True
            },
            "vehicle_no": {
                "anchors": ["Registration No.", "Vehicle No", "Regn. No"],
            },
            "idv": {
                "anchors": ["IDV", "Insured Declared Value"],
                "regex": r"([0-9,]+)"
            }
        }
    },
    {
        "id": "liberty",
        "name": "Liberty General Insurance",
        "ins_co_code": "LIBERTY",
        "keywords": ["liberty general insurance", "liberty videokon", "liberty"],
        "fields": {
            "policy_no": {
                "anchors": ["Policy Number", "Policy No"],
                "regex": r"([0-9]{15,})"
            },
            "claim_no": {
                "anchors": ["Claim Number", "Claim No"],
                "regex": r"([0-9]{15,})"
            },
            "insured_name": {
                "anchors": ["Insured Name", "Customer Name", "Name of Insured"],
            },
            "date_of_loss": {
                "anchors": ["Date of Loss", "Date of Accident"],
                "is_date": True
            },
            "policy_start_date": {
                "anchors": ["Policy Start Date", "Period of Insurance From"],
                "is_date": True
            },
            "vehicle_no": {
                "anchors": ["Registration No", "Vehicle Registration Number"],
            }
        }
    },
    {
        "id": "sompo",
        "name": "Universal Sompo",
        "ins_co_code": "SOMPO",
        "keywords": ["universal sompo", "sompo"],
        "fields": {
            "policy_no": {
                "anchors": ["Policy Number", "Policy No"],
            },
            "claim_no": {
                "anchors": ["Claim Number", "Claim No"],
            },
            "insured_name": {
                "anchors": ["Insured Name", "Name of Insured"],
            },
            "date_of_loss": {
                "anchors": ["Date of Loss", "Date of Accident"],
                "is_date": True
            },
            "policy_start_date": {
                "anchors": ["Policy Start Date", "Inception Date"],
                "is_date": True
            }
        }
    }
]


def clean_text(text):
    """Normalize extracted string, remove duplicate spaces."""
    if not text:
        return ""
    text = text.replace("\r", " ").replace("\n", " ").strip()
    return re.sub(r"\s+", " ", text)


def normalize_date(date_str):
    """Attempt to parse and standardize date formats to DD-Mmm-YY (e.g. 20-Oct-26)."""
    if not date_str:
        return ""
    date_str = date_str.strip()
    # Try various date formats common in Indian insurance forms
    formats = [
        "%d-%b-%Y", "%d-%b-%y", "%d-%B-%Y",
        "%d/%m/%Y", "%d/%m/%y",
        "%d.%m.%Y", "%d.%m.%y",
        "%Y-%m-%d", "%d-%m-%Y"
    ]
    # Remove time if attached (e.g. 20-OCT-2026 00:00:00)
    cleaned_date = re.split(r"\s+", date_str)[0].strip()
    
    for fmt in formats:
        try:
            dt = datetime.strptime(cleaned_date, fmt)
            return dt.strftime("%d-%b-%y")  # Matches e.g. 20-Oct-26
        except ValueError:
            continue
    return cleaned_date


def detect_template(full_text, templates=None):
    """Detect insurer template from full PDF text."""
    if templates is None:
        templates = DEFAULT_TEMPLATES
        
    lower_text = full_text.lower()
    for tmpl in templates:
        for kw in tmpl.get("keywords", []):
            if kw.lower() in lower_text:
                return tmpl
    # If no specific match, return a generic template
    return {
        "id": "generic",
        "name": "Generic Insurance",
        "ins_co_code": "OTHER",
        "keywords": [],
        "fields": DEFAULT_TEMPLATES[0]["fields"]
    }


def extract_from_tables(tables, field_config):
    """
    Search key-value pairs within tables (e.g. Claim Summary Sheet 2-column or 4-column tables).
    """
    anchors = field_config.get("anchors", [])
    for table in tables:
        if not table:
            continue
        for row in table:
            if not row:
                continue
            for idx, cell in enumerate(row):
                if not cell:
                    continue
                cell_clean = clean_text(cell).lower()
                for anchor in anchors:
                    anchor_clean = anchor.lower().replace("\\", "")
                    if anchor_clean in cell_clean.replace("\\", ""):
                        # Value should be in the next cell of the same row
                        if idx + 1 < len(row):
                            val_raw = row[idx + 1]
                            val = clean_text(val_raw) if val_raw else ""
                            reg = field_config.get("regex")
                            if reg and val:
                                m = re.search(reg, val, re.IGNORECASE)
                                if m:
                                    return m.group(1) if m.groups() else m.group(0)
                            return val

                        # Or in cell content after colon e.g. "Policy No: 12345"
                        parts = cell.split(":")
                        if len(parts) > 1:
                            val = clean_text(parts[1])
                            reg = field_config.get("regex")
                            if reg and val:
                                m = re.search(reg, val, re.IGNORECASE)
                                if m:
                                    return m.group(1) if m.groups() else m.group(0)
                            return val
    return None


def extract_from_raw_text(full_text, field_config):
    """
    Search lines and regex in raw text as fallback.
    """
    anchors = field_config.get("anchors", [])
    lines = full_text.splitlines()

    for anchor in anchors:
        # Regex anchor match in single line
        pattern = re.escape(anchor) + r"[\s:|\\]+([^\n\r]+)"
        match = re.search(pattern, full_text, re.IGNORECASE)
        if match:
            val = clean_text(match.group(1))
            reg = field_config.get("regex")
            if reg:
                m = re.search(reg, val, re.IGNORECASE)
                if m:
                    return m.group(1) if m.groups() else m.group(0)
            return val

        # Multi-line: check if anchor is on one line and value on next line
        for i, line in enumerate(lines):
            if anchor.lower() in line.lower():
                if i + 1 < len(lines):
                    next_val = clean_text(lines[i + 1])
                    if next_val:
                        reg = field_config.get("regex")
                        if reg:
                            m = re.search(reg, next_val, re.IGNORECASE)
                            if m:
                                return m.group(1) if m.groups() else m.group(0)
                        return next_val
    return None


def extract_data_from_pdf(pdf_path, templates=None):
    """
    Main extraction function for a given PDF file.
    Returns dictionary of extracted fields ready for Excel mapping.
    """
    if not os.path.exists(pdf_path):
        raise FileNotFoundError(f"PDF not found: {pdf_path}")

    all_text = ""
    all_tables = []

    with pdfplumber.open(pdf_path) as pdf:
        for page in pdf.pages:
            text = page.extract_text() or ""
            all_text += text + "\n"
            tables = page.extract_tables() or []
            all_tables.extend(tables)

    # 1. Determine insurer template
    template = detect_template(all_text, templates)
    
    extracted = {
        "filename": os.path.basename(pdf_path),
        "ins_co": template.get("ins_co_code", "OTHER"),
        "template_name": template.get("name", "Generic"),
        "ref_no": "",
        "proximate_days": "",
        "tat": "",
        "assign_to": "",
        "driver_relation": "",
        "ncb": "",
        "idv": "",
        "policy_no": "",
        "policy_start_date": "",
        "date_of_loss": "",
        "claim_no": "",
        "date_of_survey": datetime.now().strftime("%d-%b-%y"), # Defaults to today's date
        "insured_name": "",
        "vehicle_no": "",
        "make_model": "",
        "workshop": "",
        "survey_done_by": "",
        "advisor_name": ""
    }

    fields_config = template.get("fields", {})

    for field_key, cfg in fields_config.items():
        # First try table extraction (very accurate for tabular claims like Bajaj)
        val = extract_from_tables(all_tables, cfg)
        if val is None:
            # Fallback to text line search
            val = extract_from_raw_text(all_text, cfg)

        if val:
            val = clean_text(val)
            if cfg.get("is_date"):
                val = normalize_date(val)
            extracted[field_key] = val

    # Combine Make & Model if separate
    if not extracted.get("make_model"):
        make = extracted.get("vehicle_make", "")
        model = extracted.get("vehicle_model", "")
        extracted["make_model"] = f"{make} {model}".strip()

    # Clean Workshop / Garage name if extracted from address block
    if extracted.get("workshop"):
        raw_w = extracted["workshop"]
        clean_w = re.split(r"[-,\/]", raw_w)[0].strip()
        if len(clean_w) >= 3:
            extracted["workshop"] = clean_w

    return extracted
