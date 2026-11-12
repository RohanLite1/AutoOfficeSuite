"""
Estimate Parts Handler Module.
Extracts parts line items & amounts from bill estimate PDFs (BMW, Maruti, Hyundai, Tata, Mahindra, Toyota, Honda, Audatex, etc.)
and writes them directly into the PARTS COMPUTATION table of Master Assessment Excels (.xlsx, .xlsm, .xls).
Extracts Taxable Amounts (net value before GST) for parts and labour.
Engineered to handle small estimates as well as large estimates with hundreds of parts across multiple pages.
"""

import os
import re
import shutil
from datetime import datetime
import pdfplumber

# In-memory cache for inspect_estimate_excel: (norm_path, sheet_name) -> (mtime, size, info_dict)
_ESTIMATE_INFO_CACHE = {}


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


def invalidate_estimate_cache(file_path=None):
    """Invalidate cache for a specific file or all files."""
    global _ESTIMATE_INFO_CACHE
    if file_path:
        norm = os.path.normpath(os.path.abspath(file_path)).lower()
        keys_to_del = [k for k in list(_ESTIMATE_INFO_CACHE.keys()) if k[0].lower() == norm]
        for k in keys_to_del:
            _ESTIMATE_INFO_CACHE.pop(k, None)
    else:
        _ESTIMATE_INFO_CACHE.clear()


def clean_num(val):
    """Parses numeric amount from string, cleaning commas, currency signs, etc."""
    if val is None:
        return 0.0
    val_s = str(val).replace(",", "").replace("₹", "").replace("Rs.", "").replace("/-", "").strip()
    m = re.search(r"[-+]?\d*\.?\d+", val_s)
    if m:
        try:
            return float(m.group(0))
        except ValueError:
            return 0.0
    return 0.0


def is_labour_item(desc, uom="", hsn=""):
    """
    Determines if a line item is Labour/Service vs a physical Part.
    Indian GST rules: HSN starting with '99' is service/labour.
    """
    desc_u = (desc or "").upper()
    uom_u = (uom or "").upper()
    hsn_s = str(hsn or "").strip()

    if hsn_s.startswith("99"):
        return True

    if any(u in uom_u for u in ["FRU", "HRS", "HOUR", "MIN"]):
        return True

    # Physical components that might have words like "kit" or "cleaner"
    if any(k in desc_u for k in ["KIT", "CLEANER", "FLUID", "OIL", "FILTER", "PLUG", "TAPE", "PAD", "FASTENER", "NUT", "BOLT", "SCREW", "BULB", "GAS", "SEALANT", "GEL"]):
        return False

    labour_keywords = [
        "LABOUR", "LABOR", "JOB TIME", "DENTING", "PAINTING", "R & R", "R&R", "R / BODY",
        "REMOVAL", "REFITTING", "OVERHAUL", "DISMANTLE", "ALIGNMENT", "BALANCING",
        "WASHING", "POLISHING", "SCANNING", "CHARGES", "LABOUR CHARGES", "SERVICE CHARGE", "TOWING"
    ]
    return any(k in desc_u for k in labour_keywords)


FOOTER_STOP_KEYWORDS = [
    "company office address", "terms and conditions", "customer signature", "authorised signatory",
    "authorized signatory", "sum labor", "sum parts", "total amount", "amount to pay",
    "cgst total", "sgst total", "igst total", "customer data usage consent", "i have taken the delivery",
    "grand total", "sub total parts", "total labour charges", "amount in words"
]


def map_table_header(row):
    """Maps columns from a table header row using robust regular expressions."""
    cols = {}
    for idx, cell in enumerate(row):
        c_u = str(cell or "").lower().strip()
        if re.search(r'\b(s\.?\s*no|sl\.?\s*no|item\s*no|pos|sr\.?\s*no|es\.?)\b|^s[\.\s]*no$', c_u):
            cols['sno'] = idx
        elif re.search(r'\b(part\s*no|part\s*#|part\s*code|item\s*code|code|catalogue)\b', c_u):
            cols['code'] = idx
        elif re.search(r'\b(desc|description|particular|particulars|part\s*name|item\s*name|details|nomenclature)\b', c_u):
            cols['desc'] = idx
        elif re.search(r'\b(taxable\s*amount|taxable\s*amt|taxable\s*value|taxable)\b', c_u):
            cols['taxable_amt'] = idx
        elif re.search(r'\b(rate|unit\s*price|price|mrp|unit\s*cost)\b', c_u):
            cols['rate'] = idx
        elif re.search(r'\b(uom|unit\s*of\s*measure|u\.o\.m|^units?$)\b', c_u):
            cols['uom'] = idx
        elif re.search(r'\b(qty|quantity|qnty)\b', c_u):
            cols['qty'] = idx
        elif re.search(r'\b(amount|total|net|value|sum)\b', c_u):
            cols['amount'] = idx
        elif re.search(r'\b(hsn|sac)\b', c_u):
            cols['hsn'] = idx
    return cols


def extract_toyota_format(pdf):
    """
    Extracts parts and labour from Toyota UTrust and dealer invoices with
    borderless columnar layouts (SI.No, Job Code, Description, HSN, QTY, UOM, Unit Price, Taxable Value, etc.)
    and distinct 'Labour charges' and 'Parts charges' sections.
    """
    has_toyota = False
    for p in pdf.pages[:2]:
        txt = (p.extract_text() or "").lower()
        if "labour charges" in txt and "taxable value" in txt:
            has_toyota = True
            break
        if "job code / part" in txt or ("bpe26" in txt and "parts charges" in txt):
            has_toyota = True
            break
            
    if not has_toyota:
        return None
        
    FOOTER_KEYWORDS = [
        "terms & conditions", "gst for", "irn no", "ask for smile", "charges labour", 
        "charges parts", "total", "g.total", "rounding", "two lakhs", "#email", "#we accept",
        "authorised by", "cashier", "central gst", "state gst"
    ]
    
    items = []
    current_section = "Parts"
    
    for page in pdf.pages:
        words = page.extract_words()
        words = sorted(words, key=lambda w: (w["top"], w["x0"]))
        
        lines = []
        curr_line = []
        curr_top = None
        for w in words:
            if w["top"] < 120 or w["top"] > 765:
                continue
            if curr_top is None or abs(w["top"] - curr_top) <= 3.5:
                curr_line.append(w)
                if curr_top is None:
                    curr_top = w["top"]
            else:
                lines.append((curr_top, curr_line))
                curr_line = [w]
                curr_top = w["top"]
        if curr_line:
            lines.append((curr_top, curr_line))
            
        page_items = []
        stop_page = False
        
        for y, line_words in lines:
            if stop_page:
                break
            line_text = " ".join([w["text"] for w in line_words]).strip()
            line_lower = line_text.lower()
            
            if "labour charges" in line_lower:
                current_section = "Labour"
                continue
            elif "parts charges" in line_lower:
                current_section = "Parts"
                continue
                
            if any(kw in line_lower for kw in FOOTER_KEYWORDS):
                stop_page = True
                break
                
            sno_words = [w for w in line_words if w["x0"] < 32 and re.match(r"^\d{1,4}$", w["text"])]
            code_words = [w for w in line_words if 32 <= w["x0"] < 88 and w not in sno_words]
            
            if sno_words and code_words:
                sno = sno_words[0]["text"]
                code = " ".join([w["text"] for w in code_words])
                
                taxable_words = [w for w in line_words if 335 <= w["x0"] < 388]
                hsn_words = [w for w in line_words if 150 <= w["x0"] < 192]
                qty_words = [w for w in line_words if 192 <= w["x0"] < 214]
                rate_words = [w for w in line_words if 235 <= w["x0"] < 278]
                desc_words = [w for w in line_words if 85 <= w["x0"] < 155]
                
                taxable_amt = clean_num(taxable_words[0]["text"]) if taxable_words else 0.0
                hsn = hsn_words[0]["text"] if hsn_words else ""
                qty = qty_words[0]["text"] if qty_words else "1"
                rate = clean_num(rate_words[0]["text"]) if rate_words else taxable_amt
                desc = " ".join([w["text"] for w in desc_words])
                
                page_items.append({
                    "sno": int(sno),
                    "code": code,
                    "desc_parts": [desc] if desc else [],
                    "hsn": hsn,
                    "qty": qty,
                    "unit_price": rate,
                    "amount": taxable_amt,
                    "section": current_section
                })
            elif page_items and not stop_page:
                continuation_words = [w for w in line_words if 85 <= w["x0"] < 155]
                if continuation_words:
                    cont_text = " ".join([w["text"] for w in continuation_words])
                    page_items[-1]["desc_parts"].append(cont_text)
                    
        for it in page_items:
            full_desc = " ".join([p for p in it["desc_parts"] if p]).strip()
            full_desc = re.sub(r"\s+", " ", full_desc).strip()
            is_labour = (it["section"] == "Labour") or is_labour_item(full_desc, "", it["hsn"])
            items.append({
                "code": it["code"],
                "particular": full_desc,
                "hsn": it["hsn"],
                "qty": it["qty"],
                "unit_price": it["unit_price"],
                "amount": it["amount"],
                "item_type": "Labour" if is_labour else "Part",
                "selected": True
            })
            
    return items


def extract_toyota_claims_format(pdf):
    """
    Extracts parts and labour from Toyota Claims Portal estimates (Globe Toyota)
    where items start with 'CLM ' and have 'oApprove / oReject' buttons.
    Sections are marked by 'Labour' and 'Parts' headers.
    """
    full_text = ""
    for p in pdf.pages[:2]:
        full_text += p.extract_text() or ""
    if "CLM " not in full_text and "oApprove" not in full_text:
        return []
    
    parts = []
    labours = []
    current_section = "Labour"
    
    for page in pdf.pages:
        text = page.extract_text() or ""
        lines = text.split("\n")
        i = 0
        while i < len(lines):
            line = lines[i].strip()
            if line == "Parts":
                current_section = "Parts"
                i += 1
                continue
            elif line == "Labour":
                current_section = "Labour"
                i += 1
                continue
            
            if line.startswith("CLM "):
                clean_line = re.sub(r'oApprove.*', '', line).strip()
                clean_line = re.sub(r'oReject.*', '', clean_line).strip()
                
                # Check for continuation lines (e.g. 'NO.5', 'SUSPENSIO')
                next_line = lines[i+1].strip() if i+1 < len(lines) else ""
                if next_line and not next_line.startswith("CLM ") and not any(k in next_line for k in ["Parts", "Labour", "Estimate", "Grand Total", "Tax", "oApprove", "oReject", "Total", "Dealer Comment", "Insurance Comment", "Page"]):
                    clean_line = clean_line + " " + next_line
                    i += 1
                
                if current_section == "Labour":
                    m = re.match(r'^CLM\s+([A-Z0-9\-\/]{4,20})\s+(.+?)\s+([\d,]+\.\d{2})', clean_line)
                    if m:
                        code, desc, cost = m.groups()
                        amt = clean_num(cost)
                        labours.append({
                            "code": code.strip(),
                            "particular": desc.strip(),
                            "hsn": "998729",
                            "qty": "1",
                            "unit_price": amt,
                            "amount": amt,
                            "item_type": "Labour",
                            "selected": True
                        })
                else:
                    m = re.match(r'^CLM\s+([A-Z0-9\-\/]{6,25})\s+(.+?)\s+(\d+(?:\.\d+)?)\s+([\d,]+\.\d{2})\s+([\d,]+\.\d{2})', clean_line)
                    if m:
                        code, desc, qty, unit_cost, total_cost = m.groups()
                        amt = clean_num(total_cost)
                        rate = clean_num(unit_cost)
                        parts.append({
                            "code": code.strip(),
                            "particular": desc.strip(),
                            "hsn": "",
                            "qty": qty.strip(),
                            "unit_price": rate,
                            "amount": amt,
                            "item_type": "Part",
                            "selected": True
                        })
            i += 1
            
    return labours + parts


def extract_toyota_tserv_format(pdf):
    """
    Extracts parts and labour from Toyota T-Serv / Dealer estimates (IJM Toyota Fortuner)
    which feature a right checklist sidebar and multi-page tables across multiple pages.
    """
    full_text = ""
    for p in pdf.pages[:3]:
        full_text += p.extract_text() or ""
    if ("TOYOTA" not in full_text.upper()) or ("Part Charges" not in full_text and "Labour Charges" not in full_text) or ("[" not in full_text):
        return []
    
    parts = []
    labours = []
    current_section = "Header"
    
    for page in pdf.pages:
        text = page.extract_text() or ""
        lines = text.split("\n")
        for line in lines:
            line_s = line.strip()
            if "Labour Charges" in line_s:
                current_section = "Labour"
                continue
            elif "Part Charges" in line_s:
                current_section = "Parts"
                continue
            
            # Clean right sidebar noise
            clean = re.sub(r'\[\s*\].*', '', line_s)
            clean = re.sub(r'Id Card Driving License.*', '', clean)
            clean = re.sub(r'Claim Documents.*', '', clean)
            clean = re.sub(r'I agree to.*', '', clean)
            clean = re.sub(r'Terms & Condition.*', '', clean)
            clean = re.sub(r'Explanation of estimate.*', '', clean)
            clean = re.sub(r'I Certify That.*', '', clean)
            clean = re.sub(r'Work has been done.*', '', clean)
            clean = re.sub(r'Job details & charges.*', '', clean)
            clean = re.sub(r'Vehicle in good condition.*', '', clean)
            clean = re.sub(r'Date\s*:\s*Customer\s*Signature.*', '', clean, flags=re.I)
            clean = re.sub(r'Customer\s*Signature.*', '', clean, flags=re.I)
            clean = clean.strip()
            
            if current_section == "Labour":
                m = re.match(r'^([A-Z0-9]{5,10})\s+(.+?)\s+(\d{6})\s+([\d,]+\.\d{2})$', clean)
                if m:
                    code, desc, hsn, amt = m.groups()
                    amt_f = clean_num(amt)
                    labours.append({
                        "code": code.strip(),
                        "particular": desc.strip(),
                        "hsn": hsn.strip(),
                        "qty": "1",
                        "unit_price": amt_f,
                        "amount": amt_f,
                        "item_type": "Labour",
                        "selected": True
                    })
            elif current_section == "Parts":
                m = re.match(r'^([A-Z0-9\-\/]{6,25})\s+(.+?)\s+(\d{6,8})\s+(\d+(?:\.\d+)?)\s+([\d,]+\.\d{2})$', clean)
                if m:
                    code, desc, hsn, qty, amt = m.groups()
                    amt_f = clean_num(amt)
                    parts.append({
                        "code": code.strip(),
                        "particular": desc.strip(),
                        "hsn": hsn.strip(),
                        "qty": qty.strip(),
                        "unit_price": amt_f,
                        "amount": amt_f,
                        "item_type": "Part",
                        "selected": True
                    })
                    
    return labours + parts


def extract_bmw_dms_format(pdf):
    """
    Extracts parts and labour from BMW Dealer DMS estimates (Krishna Automobiles BMW)
    Layout: Description HSN Qty UoM Unit Price Disc. SGST CGST IGST Amount
    Parts have UoM 'PCS' / 'SET'; Labour has UoM 'FRU' / HSN '998729'.
    """
    full_text = ""
    for p in pdf.pages[:2]:
        full_text += p.extract_text() or ""
    if ("BMW" not in full_text.upper() and "Krishna Automobiles" not in full_text) or ("Sum Parts" not in full_text and "FRU" not in full_text):
        return []
        
    parts = []
    labours = []
    
    for page in pdf.pages:
        text = page.extract_text() or ""
        for line in text.split("\n"):
            line_s = line.strip()
            m = re.match(r'^(.+?)\s+(\d{6,8})\s+(\d+)\s+([A-Za-z]+)\s+([\d,.]+)\s+(?:[\d.]+%?\s+)?([\d,.]+)$', line_s)
            if m:
                desc, hsn, qty, uom, rate, amt = m.groups()
                amt_f = clean_num(amt)
                rate_f = clean_num(rate)
                is_labour = (uom.upper() == "FRU") or (hsn.startswith("99"))
                item = {
                    "code": hsn.strip(),
                    "particular": desc.strip(),
                    "hsn": hsn.strip(),
                    "qty": f"{qty} {uom}".strip(),
                    "unit_price": rate_f,
                    "amount": amt_f,
                    "item_type": "Labour" if is_labour else "Part",
                    "selected": True
                }
                if is_labour:
                    labours.append(item)
                else:
                    parts.append(item)
                    
    return labours + parts


def extract_ashok_leyland_format(pdf):
    """
    Extracts parts from Commercial Vehicle / Ashok Leyland estimates (Grover Motors)
    Layout: S.NO CODE HSN UOM QTY MRP PRICE DISC. TAXABLE CGST SGST AMOUNT
    with multi-line description continuation lines.
    """
    full_text = ""
    for p in pdf.pages[:2]:
        full_text += p.extract_text() or ""
    if "GROVER MOTORS" not in full_text.upper() and "SPARES BILL" not in full_text:
        return []
        
    parts = []
    
    for page in pdf.pages:
        text = page.extract_text() or ""
        lines = text.split("\n")
        i = 0
        while i < len(lines):
            line = lines[i].strip()
            m = re.match(r'^(\d{1,3})\s+([A-Za-z0-9_\-\(\)\/\.]{3,25})\s+(\d{6,8})\s+([A-Za-z]{1,4})\s+(\d+(?:\.\d+)?)\s+([\d,]+\.\d{2})\s+([\d,]+\.\d{2})\s+(?:[\d\.]+\s*%|[\d,]+\.\d{2})?\s+([\d,]+\.\d{2})\s+', line)
            if m:
                sno, code, hsn, uom, qty, mrp, price, taxable_amt = m.groups()
                amt = clean_num(taxable_amt)
                rate = clean_num(price)
                
                desc_parts = []
                j = i + 1
                while j < len(lines):
                    next_l = lines[j].strip()
                    if not next_l or re.match(r'^\d{1,3}\s+[A-Za-z0-9_\-\(\)\/\.]{3,25}\s+\d{6,8}', next_l) or "Subtotal" in next_l or "Printed on" in next_l or "INVOICE NO" in next_l or "AMOUNT IN WORDS" in next_l:
                        break
                    desc_text = re.sub(r'(?:\s+[\d,]+\.\d{2})+$', '', next_l).strip()
                    if desc_text and not any(k in desc_text for k in ["S.N CODE", "SPARES BILL", "CUSTOMER", "VEHICLE"]):
                        desc_parts.append(desc_text)
                    j += 1
                
                full_desc = " ".join(desc_parts).strip()
                parts.append({
                    "srl": sno.strip(),
                    "code": code.strip(),
                    "particular": full_desc if full_desc else code.strip(),
                    "hsn": hsn.strip(),
                    "qty": f"{qty} {uom}".strip(),
                    "unit_price": rate,
                    "amount": amt,
                    "item_type": "Part",
                    "selected": True
                })
                i = j - 1
            i += 1
            
    return parts


def extract_maruti_sandhu_format(pdf):
    """
    Extracts parts and labour from Maruti Suzuki dealership estimates (Sandhu Automobiles / Service Estimate)
    Columns: Srl | Part Number | Description | Tax | Qty. | Rate | Taxable Amount | Tax Paid Amount | Labor Amount
    Extracts Taxable Amount (Col 6), ignoring the zero-filled Tax Paid Amount (Col 7).
    """
    full_text = ""
    for p in pdf.pages[:2]:
        full_text += p.extract_text() or ""
    if "Service Estimate" not in full_text and "Tax Paid" not in full_text:
        return []
    if "Taxable" not in full_text:
        return []

    parts = []
    labours = []
    current_sec = "Parts"

    for page in pdf.pages:
        text = page.extract_text() or ""
        lines = text.split("\n")
        for line in lines:
            line_s = line.strip()
            if line_s == "Labour":
                current_sec = "Labour"
                continue
            if line_s == "Parts":
                current_sec = "Parts"
                continue
            if line_s.startswith("Total :") or line_s.startswith("Sub Total") or "Net Estimate" in line_s or "Estimated Days" in line_s or "SGST @" in line_s or "CGST @" in line_s:
                continue

            # Maruti 4-number Part Line: Srl | PartNo | Description | Qty | Rate | Taxable Amt | Tax Paid Amt
            # e.g.: 1 17100M74L50 FAN ASSY, ENGINE COOLING 1.000 5,220.00 5,220.00 0.00
            # e.g.: 25 72390M74L00 CROSSMEMBER, FRONT LOWER 1.000 1,299.50 1,299.50 0.00
            m_part = re.match(
                r"^(\d{1,4})\s+([A-Z0-9\-\/]{6,25})\s+(.+?)\s+(\d+(?:\.\d+)?)\s+([\d,]+\.\d{2})\s+([\d,]+\.\d{2})\s+([\d,]+\.\d{2})$",
                line_s
            )
            if m_part:
                srl, code, desc, qty, rate, taxable_amt, tax_paid = m_part.groups()
                amt = clean_num(taxable_amt)
                rate_f = clean_num(rate)
                qty_clean = f"{clean_num(qty):.0f}" if clean_num(qty).is_integer() else str(clean_num(qty))
                parts.append({
                    "srl": srl.strip(),
                    "code": code.strip(),
                    "particular": desc.strip(),
                    "hsn": "",
                    "qty": qty_clean,
                    "unit_price": rate_f,
                    "amount": amt,
                    "item_type": "Part",
                    "selected": True
                })
                continue

            # Maruti Labour Line: Srl | Code | Description | Amount
            # e.g.: 1 ZF9993 PAINTING CHARGES 30,000.00
            # e.g.: 2 ZF9992 DENTING CHARGES 10,000.00
            m_lab = re.match(r"^(\d{1,4})\s+([A-Z0-9]{4,10})\s+(.+?)\s+([\d,]+\.\d{2})$", line_s)
            if m_lab and current_sec == "Labour":
                srl, code, desc, amt_s = m_lab.groups()
                amt = clean_num(amt_s)
                labours.append({
                    "srl": srl.strip(),
                    "code": code.strip(),
                    "particular": desc.strip(),
                    "hsn": "998729",
                    "qty": "1",
                    "unit_price": amt,
                    "amount": amt,
                    "item_type": "Labour",
                    "selected": True
                })

    return parts + labours


def extract_printed_estimate_summary(pdf):
    """
    Extracts printed parts total, labour total, and grand total from the estimate summary block.
    Used for mathematical checksum reconciliation (Zero-Miss Guarantee).
    """
    summary = {
        "parts_total": None,
        "labour_total": None,
        "grand_total": None
    }
    full_text = "\n".join([(p.extract_text() or "") for p in pdf.pages])
    
    # 1. Parts Total
    m_p = re.findall(r'(?:Sum\s+Parts|Parts\s+Total|Total\s+Parts|Part\s*Amt\s*:)[:\s]*([\d,]+\.\d{2})', full_text, re.I)
    if not m_p:
        # Maruti 3-column total line: Total : <PartsTotal> <TaxPaid> <LabourTotal>
        # e.g.: Total : 1,15,547.50 0.00 40,000.00
        m_maruti_tot = re.findall(r'Total\s*:\s*([\d,]+\.\d{2})\s+[\d,]+\.\d{2}\s+([\d,]+\.\d{2})', full_text)
        if m_maruti_tot:
            m_p = [m_maruti_tot[-1][0]]
            if summary["labour_total"] is None:
                summary["labour_total"] = clean_num(m_maruti_tot[-1][1])
    if not m_p:
        m_p = re.findall(r'(?:^|\n)\s*Parts\s+([\d,]+\.\d{2})\s+0\.00', full_text, re.I)
    if not m_p:
        m_p = re.findall(r'(?:^|\n)\s*Parts\s+([\d,]+\.\d{2})', full_text, re.I)
    if not m_p:
        m_p = re.findall(r'Subtotal\s+[\d,\.]+\s+([\d,]+\.\d{2})', full_text, re.I)
    if m_p:
        summary["parts_total"] = clean_num(m_p[-1])
        
    # 2. Labour Total
    if summary["labour_total"] is None:
        m_l = re.findall(r'(?:Sum\s+Labou?r|Labou?r\s+Total|Total\s+Labou?r|Labou?r\s*Amt\s*:)[:\s]*([\d,]+\.\d{2})', full_text, re.I)
        if not m_l:
            m_l_bmw = re.findall(r'(\d{5,8}\.\d{2})\s+\d{6,8}\.\d{2}\s+0\.00', full_text)
            if m_l_bmw:
                m_l = m_l_bmw
        if not m_l:
            m_l = re.findall(r'(?:^|\n)\s*(?:Labour|Item\s*Amount\s*Discount\s*\n)\s*([\d,]+\.\d{2})', full_text, re.I)
        if m_l:
            summary["labour_total"] = clean_num(m_l[-1])
        
    # 3. Grand Total
    m_g = re.findall(r'(?:Grand\s*Total|Total\s*Amt\s*:|Net\s*Estimate\s*Amounts?\s*:|Total\s*Amount\s*Incl\.\s*GST)[:\s]*([\d,]+\.\d{2})', full_text, re.I)
    if m_g:
        summary["grand_total"] = clean_num(m_g[-1])
        
    return summary


def parse_estimate_pdf(pdf_path):
    """
    Parses a dealer bill estimate PDF (BMW, Maruti, Hyundai, Tata, Mahindra, Toyota, Honda, Audatex, etc.)
    Returns header metadata and full list of extracted line items (Parts and Labour).
    Extracts Taxable Amounts (net value before GST) as requested for assessment.
    Capable of extracting hundreds of parts across multi-page estimates.
    """
    if not os.path.exists(pdf_path):
        raise FileNotFoundError(f"PDF not found: {pdf_path}")

    header_info = {
        "invoice_no": "",
        "invoice_date": "",
        "registration_no": "",
        "make": "",
        "model": "",
        "service_advisor": "",
        "dealer_name": ""
    }
    items = []

    with pdfplumber.open(pdf_path) as pdf:
        # Header extraction from initial pages
        for p_idx in range(min(2, len(pdf.pages))):
            text = pdf.pages[p_idx].extract_text() or ""
            lines = text.splitlines()

            if not header_info["registration_no"]:
                m_reg = re.search(r'\b([A-Z]{2}\s*\d{1,2}\s*[A-Z]{1,3}\s*\d{4})\b', text)
                if m_reg:
                    header_info['registration_no'] = m_reg.group(1).replace(' ', '')

            if not header_info["invoice_no"]:
                m_inv = re.search(r'Invoice\s*No\.?\s*[:\.\s]*([A-Z0-9\-\/]{4,25})', text, re.I)
                if m_inv:
                    header_info['invoice_no'] = m_inv.group(1).strip()
                else:
                    m_jc = re.search(r'(?:Bill|Estimate|Job\s*Card)\s*No\.?\s*[:\.\s]*([A-Z0-9\-\/]{4,25})', text, re.I)
                    if m_jc:
                        header_info['invoice_no'] = m_jc.group(1).strip()

            if not header_info["invoice_date"]:
                m_date = re.search(r'(?:Date|Invoice\s*Date|Estimate\s*Date)[:\.\s]*([0-9]{1,2}[\/\-\.][0-9]{1,2}[\/\-\.][0-9]{2,4}|[0-9]{1,2}-[A-Za-z]{3}-[0-9]{2,4})', text, re.I)
                if m_date:
                    header_info['invoice_date'] = m_date.group(1).strip()

            if not header_info["service_advisor"]:
                m_adv = re.search(r'(?:SA\s*Name|Service\s*Advisor)[:\.\s]*([A-Za-z\s]+?)(?:SA\(M\)|Phone|Ph|\n|$)', text, re.I)
                if m_adv:
                    header_info['service_advisor'] = m_adv.group(1).strip()

            if not header_info["model"]:
                m_model = re.search(r'Model\s*[:\s]*([A-Za-z0-9\s\-]+?)(?:\s+Sold by|\s+Engine|\s+Chassis|\n|$)', text, re.I)
                if m_model:
                    header_info['model'] = m_model.group(1).strip()

            if not header_info["dealer_name"]:
                for line in lines[:8]:
                    line_s = line.strip()
                    if any(term in line_s.lower() for term in ['automobiles', 'motors', 'bmw', 'maruti', 'hyundai', 'workshop', 'services', 'dealership', 'garage', 'tata', 'mahindra', 'toyota', 'honda']):
                        header_info['dealer_name'] = line_s.split('   ')[0].strip()
                        break
                if not header_info["dealer_name"]:
                    m_dealer = re.search(r'For\s+([A-Za-z0-9\s\.\,\-]+?(?:AUTOMOBILES|MOTORS|WORKSHOP|SERVICES|PVT\.?\s*LTD\.?))', text, re.I)
                    if m_dealer:
                        header_info['dealer_name'] = m_dealer.group(1).strip()

        # Pass 0.1: Specialized High-Precision DMS Extractors (100% Zero-Miss Architecture)
        specialized_extractors = [
            ("Toyota Claims Portal", extract_toyota_claims_format),
            ("Toyota T-Serv", extract_toyota_tserv_format),
            ("BMW DMS", extract_bmw_dms_format),
            ("Ashok Leyland Commercial", extract_ashok_leyland_format),
            ("Maruti Suzuki Service Estimate", extract_maruti_sandhu_format),
            ("Toyota UTrust Columnar", extract_toyota_format)
        ]
        for name, extractor_fn in specialized_extractors:
            spec_items = extractor_fn(pdf)
            if spec_items and len(spec_items) >= 2:
                items = spec_items
                break

        table_items = []
        if not items:
            # Pass 1: Multi-page Table Extraction (with multi-page continuation table support)
            last_cols = None
            last_table_type = None

            for p_idx, page in enumerate(pdf.pages):
                tables = page.extract_tables() or []
                for table in tables:
                    if not table or len(table) < 1:
                        continue

                    h_idx = -1
                    cols = {}
                    for r_idx, row in enumerate(table[:5]):
                        cols_candidate = map_table_header(row)
                        if "desc" in cols_candidate and ("taxable_amt" in cols_candidate or "amount" in cols_candidate or "rate" in cols_candidate):
                            h_idx = r_idx
                            cols = cols_candidate
                            break

                    if h_idx >= 0 and "desc" in cols:
                        last_cols = cols
                        first_row_text = " ".join([str(c or "").strip() for c in table[0]]).lower()
                        last_table_type = "Labour" if "labour" in first_row_text else "Part"
                        rows_to_process = table[h_idx + 1:]
                    elif last_cols and len(table[0]) >= 4 and not any(term in " ".join([str(c or "").strip() for c in table[0]]).lower() for term in ["estimate", "total", "customer", "prepared by", "terms"]):
                        # Continuation table on subsequent page without repeating header!
                        cols = last_cols
                        rows_to_process = table
                    else:
                        continue

                    amt_col = cols.get("taxable_amt") or cols.get("amount") or cols.get("rate")
                    for row in rows_to_process:
                        if not row:
                            continue
                        row_text = " ".join([str(c or "").strip() for c in row]).lower()
                        if any(term in row_text for term in ["total", "sub total", "grand total", "cgst", "sgst", "igst", "round off", "balance", "terms and cond"]):
                            continue

                        desc = str(row[cols["desc"]] or "").strip() if "desc" in cols and cols["desc"] < len(row) else ""
                        amt_raw = row[amt_col] if amt_col < len(row) else 0
                        amt = clean_num(amt_raw)

                        if desc and amt > 0:
                            code = str(row[cols["code"]] or "").strip() if "code" in cols and cols["code"] < len(row) else ""
                            qty = str(row[cols["qty"]] or "1").strip() if "qty" in cols and cols["qty"] < len(row) else "1"
                            uom = str(row[cols["uom"]] or "").strip() if "uom" in cols and cols["uom"] < len(row) else ""
                            hsn = str(row[cols["hsn"]] or "").strip() if "hsn" in cols and cols["hsn"] < len(row) else ""
                            rate = clean_num(row[cols["rate"]]) if "rate" in cols and cols["rate"] < len(row) else amt

                            is_labour = (last_table_type == "Labour") or is_labour_item(desc, uom, hsn)
                            table_items.append({
                                "code": code,
                                "particular": desc,
                                "hsn": hsn,
                                "qty": f"{qty} {uom}".strip(),
                                "unit_price": rate or amt,
                                "amount": amt,
                                "item_type": "Labour" if is_labour else "Part",
                                "selected": True
                            })

        # Pass 1.5: If few items found, try Text-based Table Extraction
        if not items and len(table_items) < 2:
            table_items_text = []
            for p_idx, page in enumerate(pdf.pages):
                tables = page.extract_tables({"vertical_strategy": "text", "horizontal_strategy": "text"}) or []
                for table in tables:
                    if not table or len(table) < 2:
                        continue
                    h_idx = -1
                    cols = {}
                    for r_idx, row in enumerate(table[:5]):
                        cols_candidate = map_table_header(row)
                        if "desc" in cols_candidate and ("taxable_amt" in cols_candidate or "amount" in cols_candidate or "rate" in cols_candidate):
                            h_idx = r_idx
                            cols = cols_candidate
                            break
                    if h_idx >= 0 and "desc" in cols and ("taxable_amt" in cols or "amount" in cols or "rate" in cols):
                        amt_col = cols.get("taxable_amt") or cols.get("amount") or cols.get("rate")
                        for row in table[h_idx + 1:]:
                            if not row:
                                continue
                            row_text = " ".join([str(c or "").strip() for c in row]).lower()
                            if any(term in row_text for term in ["total", "sub total", "grand total", "cgst", "sgst", "igst", "round off", "balance", "terms and cond"]):
                                continue
                            desc = str(row[cols["desc"]] or "").strip() if "desc" in cols and cols["desc"] < len(row) else ""
                            amt_raw = row[amt_col] if amt_col < len(row) else 0
                            amt = clean_num(amt_raw)
                            if desc and amt > 0:
                                code = str(row[cols["code"]] or "").strip() if "code" in cols and cols["code"] < len(row) else ""
                                qty = str(row[cols["qty"]] or "1").strip() if "qty" in cols and cols["qty"] < len(row) else "1"
                                uom = str(row[cols["uom"]] or "").strip() if "uom" in cols and cols["uom"] < len(row) else ""
                                hsn = str(row[cols["hsn"]] or "").strip() if "hsn" in cols and cols["hsn"] < len(row) else ""
                                rate = clean_num(row[cols["rate"]]) if "rate" in cols and cols["rate"] < len(row) else amt
                                is_labour = is_labour_item(desc, uom, hsn)
                                table_items_text.append({
                                    "code": code,
                                    "particular": desc,
                                    "hsn": hsn,
                                    "qty": f"{qty} {uom}".strip(),
                                    "unit_price": rate or amt,
                                    "amount": amt,
                                    "item_type": "Labour" if is_labour else "Part",
                                    "selected": True
                                })
            if len(table_items_text) > len(table_items):
                table_items = table_items_text

        if not items and len(table_items) >= 2:
            items = table_items
        elif not items:
            # Pass 2: High-capacity multi-pattern text streaming with continuation line aggregation
            for p_idx, page in enumerate(pdf.pages):
                text = page.extract_text() or ""
                lines = text.splitlines()
                last_item_idx_on_page = -1

                for line in lines:
                    line_s = line.strip()
                    if not line_s:
                        continue

                    line_lower = line_s.lower()

                    # Stop processing on page footer
                    if any(stop_kw in line_lower for stop_kw in FOOTER_STOP_KEYWORDS):
                        break

                    # Header/metadata skips
                    if any(term in line_lower for term in [
                        "tax invoice", "vehicle owner", "make code vin", "model registration",
                        "place of supply", "irn :", "dealer", "customer phone", "gstin:", "pan:",
                        "service order", "no. description hsn qty", "s.no part no", "demanded repairs", "gate pass"
                    ]):
                        continue

                    # Pattern 1: Maruti Suzuki DMS format with Taxable Amount:
                    # Srl | Part No | Description | Batch | HSN | Tax% | Qty | Rate | Taxable Amt | Tax Paid Amt | Labour
                    # e.g.: 1 09471M12210 BULB (12V,35W H8) AH 85122020 18% 1.000 504.23 504.23 0.00
                    # e.g.: 16 17861M68P00 BRACKET,RDTR SUPPORT UPPER AA 87089900 18% 2.000 27.11 54.22 0.00
                    m_maruti = re.match(
                        r"^(\d{1,4})\s+([A-Z0-9\-\/]{6,25})\s+(.+?)\s+([A-Z]{1,3})?\s*(\d{6,8})\s+(\d+%\s+)?(\d+(?:\.\d+)?)\s+([\d,]+\.\d{2})\s+([\d,]+\.\d{2})(?:\s+([\d,]+\.\d{2}))?",
                        line_s
                    )
                    if m_maruti:
                        srl, code, desc, batch, hsn, tax, qty, rate, taxable_amt, tax_paid = m_maruti.groups()
                        amt = clean_num(taxable_amt)  # Always use TAXABLE AMOUNT!
                        is_labour = is_labour_item(desc, "", hsn or "")
                        items.append({
                            "code": code.strip(),
                            "particular": desc.strip(),
                            "hsn": hsn.strip() if hsn else "",
                            "qty": f"{qty} NOS".strip(),
                            "unit_price": clean_num(rate),
                            "amount": amt,
                            "item_type": "Labour" if is_labour else "Part",
                            "selected": True
                        })
                        last_item_idx_on_page = len(items) - 1
                        continue

                    # Pattern 2: Maruti Labour format:
                    # e.g.: 1 ZA38L0 TOWING CHARGES 998729 1,270.00
                    # e.g.: 2 ZF9993 PAINTING CHARGES 998729 42,000.00
                    m_labour = re.match(r"^(\d{1,4})\s+([A-Z0-9]{5,10})\s+(.+?)\s+(\d{6})\s+([\d,]+\.\d{2})$", line_s)
                    if m_labour:
                        srl, code, desc, hsn, amt = m_labour.groups()
                        amt_num = clean_num(amt)
                        items.append({
                            "code": code.strip(),
                            "particular": desc.strip(),
                            "hsn": hsn.strip(),
                            "qty": "1",
                            "unit_price": amt_num,
                            "amount": amt_num,
                            "item_type": "Labour",
                            "selected": True
                        })
                        last_item_idx_on_page = len(items) - 1
                        continue

                    # Pattern 2.5: Kia / Dealer Labour with percentage and taxable value
                    # e.g.: 1 A10VABODY01LNA Body & Paint Repair Charges (Insurance) 998729 18% 1.00 13,600.00 100.00% 13,600.00
                    m_kia_labour = re.match(
                        r"^(\d{1,4})\s+([A-Z0-9]{8,25})\s+(.+?)\s+(\d{6})\s+(\d+%\s+)?(\d+(?:\.\d+)?)\s+([\d,]+\.\d{2})\s+.*?\s+([\d,]+\.\d{2})$",
                        line_s
                    )
                    if m_kia_labour:
                        srl, code, desc, hsn, tax_p, qty, rate, amt = m_kia_labour.groups()
                        amt_num = clean_num(amt) or clean_num(rate)
                        items.append({
                            "code": code.strip(),
                            "particular": desc.strip(),
                            "hsn": hsn.strip(),
                            "qty": f"{qty}".strip(),
                            "unit_price": clean_num(rate),
                            "amount": amt_num,
                            "item_type": "Labour",
                            "selected": True
                        })
                        last_item_idx_on_page = len(items) - 1
                        continue

                    # Pattern 3: Full Dealer/BMW format: Code Desc HSN Qty UoM Rate Disc DiscAmt CGST SGST Total
                    # e.g.: 83192289285 Repair flat 35069190 2 PCS 6775.00 10.0% 1355.00 9.0% 9.0% 12195.00
                    # e.g.: 5199000 Job time without allowance fo 998729 40 FRU 303.00 10.0% 1212.00 9.0% 9.0% 10908.00
                    mA = re.match(
                        r"^([A-Z0-9\-\/]{4,25})\s+(.+?)\s+(\d{4,8})\s+(\d+(?:\.\d+)?)\s+([A-Za-z]+)\s+([\d,]+\.?\d*)\s+.*?\s+([\d,]+\.\d{2})$",
                        line_s
                    )
                    if mA:
                        code, desc, hsn, qty, uom, unit_price, amt = mA.groups()
                        is_labour = is_labour_item(desc, uom, hsn)
                        items.append({
                            "code": code.strip(),
                            "particular": desc.strip(),
                            "hsn": hsn.strip(),
                            "qty": f"{qty} {uom}".strip(),
                            "unit_price": clean_num(unit_price),
                            "amount": clean_num(amt),
                            "item_type": "Labour" if is_labour else "Part",
                            "selected": True
                        })
                        last_item_idx_on_page = len(items) - 1
                        continue

                    # Pattern 4: S.No + Part Code + Description + Qty + Rate + Amount
                    # e.g.: 1 71711M68P00-799 BUMPER, FRONT (PRIMED) 1.00 3250.00 3250.00
                    # e.g.: 10 09119M16008 CLIP, BUMPER 10.00 PCS 15.00 150.00
                    mB = re.match(
                        r"^(?:\d{1,4}[\.\)]?\s+)?([A-Z0-9\-\/]{4,25})\s+(.+?)\s+(\d+(?:\.\d+)?)\s*([A-Za-z]{2,5})?\s+([\d,]+\.?\d*)\s+([\d,]+\.\d{2})$",
                        line_s
                    )
                    if mB:
                        code, desc, qty, uom, rate, amt = mB.groups()
                        desc_s = desc.strip()
                        if desc_s.lower() not in ["total", "sub total", "amount"]:
                            is_labour = is_labour_item(desc_s, uom or "", "")
                            items.append({
                                "code": code.strip(),
                                "particular": desc_s,
                                "hsn": "",
                                "qty": f"{qty} {uom or ''}".strip(),
                                "unit_price": clean_num(rate),
                                "amount": clean_num(amt),
                                "item_type": "Labour" if is_labour else "Part",
                                "selected": True
                            })
                            last_item_idx_on_page = len(items) - 1
                            continue

                    # Pattern 5: Code + Description + Qty + Amount (Simplified 4-field)
                    # e.g.: 51319478772 Frontscheibe 1 PCS 48086.32
                    mC = re.match(
                        r"^([A-Z0-9\-\/]{4,25})\s+([A-Za-z0-9\s\.\-\/\(\)&,]{3,60}?)\s+(\d+(?:\.\d+)?)\s*([A-Za-z]{2,5})?\s+([\d,]+\.\d{2})$",
                        line_s
                    )
                    if mC:
                        code, desc, qty, uom, amt = mC.groups()
                        desc_s = desc.strip()
                        if desc_s.lower() not in ["total", "sub total", "amount"]:
                            is_labour = is_labour_item(desc_s, uom or "", "")
                            items.append({
                                "code": code.strip(),
                                "particular": desc_s,
                                "hsn": "",
                                "qty": f"{qty} {uom or ''}".strip(),
                                "unit_price": clean_num(amt),
                                "amount": clean_num(amt),
                                "item_type": "Labour" if is_labour else "Part",
                                "selected": True
                            })
                            last_item_idx_on_page = len(items) - 1
                            continue

                    # Pattern 6: S.No + Description + Qty + Amount (No Part Code)
                    # e.g.: 1. FRONT BUMPER ASSY 1 NOS 4,500.00
                    # e.g.: 15 RADIATOR SUPPORT LH 1 2300.00
                    mD = re.match(
                        r"^(\d{1,4})[\.\)]?\s+([A-Za-z0-9\s\.\-\/\(\)&,]{3,60}?)\s+(\d+(?:\.\d+)?)\s*([A-Za-z]{2,5})?\s+([\d,]+\.\d{2})$",
                        line_s
                    )
                    if mD:
                        sno, desc, qty, uom, amt = mD.groups()
                        desc_s = desc.strip()
                        if desc_s.lower() not in ["total", "sub total", "amount", "page"]:
                            is_labour = is_labour_item(desc_s, uom or "", "")
                            items.append({
                                "code": f"#{sno}",
                                "particular": desc_s,
                                "hsn": "",
                                "qty": f"{qty} {uom or ''}".strip(),
                                "unit_price": clean_num(amt),
                                "amount": clean_num(amt),
                                "item_type": "Labour" if is_labour else "Part",
                                "selected": True
                            })
                            last_item_idx_on_page = len(items) - 1
                            continue

                    # Controlled continuation: only if immediately following an item on this page
                    if last_item_idx_on_page >= 0 and len(line_s) < 30 and not re.search(r'\d+\.\d{2}', line_s) and not line_s.startswith("Job-"):
                        if not any(k in line_lower for k in ["signature", "address", "phone", "email", "pan", "gstin", "order", "date", "road", "pin", "ludhiana", "delhi"]):
                            items[last_item_idx_on_page]["particular"] += f" {line_s}"

        # Extract printed summary before closing PDF
        printed_summary = extract_printed_estimate_summary(pdf)

    # Ensure every item has a clean, reliable serial number (srl)
    p_seq = 1
    l_seq = 1
    for it in items:
        if not it.get("srl") and it.get("sno"):
            it["srl"] = str(it["sno"]).strip()
        if not it.get("srl") or not str(it.get("srl")).strip():
            if it.get("item_type") == "Labour":
                it["srl"] = str(l_seq)
                l_seq += 1
            else:
                it["srl"] = str(p_seq)
                p_seq += 1
        else:
            srl_str = re.sub(r'[\.\s]', '', str(it["srl"]))
            if srl_str.isdigit():
                it["srl"] = srl_str

    parts_count = sum(1 for it in items if it["item_type"] == "Part")
    parts_sum = sum(it["amount"] for it in items if it["item_type"] == "Part")
    labour_count = sum(1 for it in items if it["item_type"] == "Labour")
    labour_sum = sum(it["amount"] for it in items if it["item_type"] == "Labour")

    p_target = printed_summary.get("parts_total")
    reconciliation_verified = False
    diff = 0.0
    if p_target and p_target > 0:
        diff = round(parts_sum - p_target, 2)
        if abs(diff) < 1.0:
            reconciliation_verified = True

    return {
        "success": True,
        "filename": os.path.basename(pdf_path),
        "header": header_info,
        "items": items,
        "total_parts": parts_count,
        "total_parts_sum": round(parts_sum, 2),
        "total_labour": labour_count,
        "total_labour_sum": round(labour_sum, 2),
        "total_estimate_sum": round(parts_sum + labour_sum, 2),
        "printed_summary": printed_summary,
        "reconciliation": {
            "verified": reconciliation_verified,
            "printed_parts_total": p_target,
            "difference": diff
        }
    }


def inspect_estimate_excel(file_path, target_sheet=None):
    """
    Inspects an Excel spreadsheet (.xlsx, .xlsm, or .xls) to locate the active PARTS COMPUTATION table.
    Scans VISIBLE rows to locate the visible table header and first visible data row (e.g. Row 171).
    Uses mtime caching for instant responses.
    """
    if not os.path.exists(file_path):
        return {"success": False, "error": f"File does not exist: {file_path}"}

    norm_path = os.path.normpath(os.path.abspath(file_path))
    try:
        stat = os.stat(norm_path)
        mtime, size = stat.st_mtime, stat.st_size
    except OSError:
        return {"success": False, "error": f"Cannot access file: {file_path}"}

    cache_key = (norm_path, str(target_sheet or "").strip())
    if cache_key in _ESTIMATE_INFO_CACHE:
        cached_mtime, cached_size, cached_info = _ESTIMATE_INFO_CACHE[cache_key]
        if cached_mtime == mtime and cached_size == size:
            return cached_info.copy()

    ext = os.path.splitext(norm_path)[1].lower()

    if ext == ".xls":
        try:
            import xlrd
            wb = xlrd.open_workbook(file_path, formatting_info=True)
            sheet_names = wb.sheet_names()

            best_sheet = ""
            best_score = -1
            best_header_r = 169
            best_start_r = 171
            best_cols = {"es": 1, "bill_no": 2, "particular": 3, "estimate": 4}

            for sname in sheet_names:
                s = wb.sheet_by_name(sname)
                if s.visibility != 0 and sname != target_sheet:
                    continue

                score = 0
                h_row = -1
                cols = {"es": 1, "bill_no": 2, "particular": 3, "estimate": 4}

                for r in range(s.nrows):
                    rinfo = s.rowinfo_map.get(r) if hasattr(s, "rowinfo_map") else None
                    if rinfo and rinfo.hidden:
                        continue

                    row_vals = [str(s.cell_value(r, c)).strip() for c in range(min(20, s.ncols))]
                    row_str = " ".join(row_vals).upper()

                    if "PARTS COMPUTATION" in row_str:
                        score += 50
                        for r2 in range(r + 1, min(r + 15, s.nrows)):
                            r2info = s.rowinfo_map.get(r2) if hasattr(s, "rowinfo_map") else None
                            if r2info and r2info.hidden:
                                continue
                            r2_vals = [str(s.cell_value(r2, c)).strip() for c in range(min(20, s.ncols))]
                            if any("PARTICULAR" in v.upper() for v in r2_vals):
                                h_row = r2 + 1
                                for c, v in enumerate(r2_vals):
                                    v_u = v.upper()
                                    if "PARTICULAR" in v_u: cols["particular"] = c + 1
                                    elif "ESTIMATE" in v_u: cols["estimate"] = c + 1
                                    elif "BILL" in v_u: cols["bill_no"] = c + 1
                                    elif "ES." in v_u or v_u == "ES" or "S.NO" in v_u: cols["es"] = c + 1
                                break
                        if h_row > 0:
                            break

                    elif any("PARTICULAR" in v.upper() for v in row_vals) and any("ESTIMATE" in v.upper() for v in row_vals):
                        score += 40
                        h_row = r + 1
                        for c, v in enumerate(row_vals):
                            v_u = v.upper()
                            if "PARTICULAR" in v_u: cols["particular"] = c + 1
                            elif "ESTIMATE" in v_u: cols["estimate"] = c + 1
                            elif "BILL" in v_u: cols["bill_no"] = c + 1
                            elif "ES." in v_u or v_u == "ES" or "S.NO" in v_u: cols["es"] = c + 1
                        break

                if sname.lower() == "sheet1":
                    score += 15
                if target_sheet and sname.strip().lower() == target_sheet.strip().lower():
                    score += 200

                if score > best_score:
                    best_score = score
                    best_sheet = sname
                    best_cols = cols
                    if h_row > 0:
                        best_header_r = h_row
                        best_start_r = h_row + 1
                        for r_data in range(h_row, min(h_row + 15, s.nrows)):
                            r_d_info = s.rowinfo_map.get(r_data) if hasattr(s, "rowinfo_map") else None
                            if r_d_info and r_d_info.hidden:
                                continue
                            sub_vals = [str(s.cell_value(r_data, c)).strip().lower() for c in range(min(15, s.ncols))]
                            if any(term in sub_vals for term in ["metal", "non metal", "glass", "labour", "tax"]):
                                continue
                            best_start_r = r_data + 1
                            break

            if not best_sheet and sheet_names:
                best_sheet = sheet_names[0]

            # Detect LABOUR COMPUTATION start row in best_sheet
            labour_start_r = 359
            if best_sheet and best_sheet in sheet_names:
                s_best = wb.sheet_by_name(best_sheet)
                for r_lab in range(s_best.nrows):
                    r_lab_vals = [str(s_best.cell_value(r_lab, c)).strip().upper() for c in range(min(15, s_best.ncols))]
                    if "LABOUR COMPUTATION" in " ".join(r_lab_vals):
                        for r_under in range(r_lab + 1, min(r_lab + 20, s_best.nrows)):
                            r_u_vals = [str(s_best.cell_value(r_under, c)).strip().upper() for c in range(min(15, s_best.ncols))]
                            if "ASSESSED AS UNDER" in " ".join(r_u_vals):
                                labour_start_r = r_under + 2
                                break
                            if re.match(r"^1(\.0)?$", r_u_vals[0]):
                                labour_start_r = r_under + 1
                                break
                        break

            res = {
                "success": True,
                "file_path": norm_path,
                "file_type": "xls",
                "sheets": sheet_names,
                "selected_sheet": best_sheet,
                "header_row": best_header_r,
                "data_start_row": best_start_r,
                "labour_start_row": labour_start_r,
                "particular_col": best_cols["particular"],
                "estimate_col": best_cols["estimate"],
                "bill_no_col": best_cols["bill_no"],
                "es_col": best_cols["es"],
                "particular_col_letter": chr(64 + best_cols["particular"]) if 1 <= best_cols["particular"] <= 26 else "C",
                "estimate_col_letter": chr(64 + best_cols["estimate"]) if 1 <= best_cols["estimate"] <= 26 else "D"
            }
            _ESTIMATE_INFO_CACHE[cache_key] = (mtime, size, res)
            return res.copy()
        except Exception as e:
            return {"success": False, "error": f"Failed reading .xls file: {str(e)}"}

    else:
        # .xlsx or .xlsm
        wb = None
        try:
            import openpyxl
            wb = openpyxl.load_workbook(norm_path, data_only=True)
            sheet_names = wb.sheetnames

            best_sheet = ""
            best_score = -1
            best_header_r = 169
            best_start_r = 171
            best_cols = {"es": 1, "bill_no": 2, "particular": 3, "estimate": 4}

            for sname in sheet_names:
                sheet = wb[sname]
                score = 0
                h_row = -1
                cols = {"es": 1, "bill_no": 2, "particular": 3, "estimate": 4}

                for r in range(1, min(sheet.max_row + 1, 600)):
                    dim = sheet.row_dimensions.get(r)
                    if dim and dim.hidden:
                        continue

                    row_vals = [str(sheet.cell(row=r, column=c).value or "").strip() for c in range(1, min(20, sheet.max_column + 1))]
                    row_str = " ".join(row_vals).upper()

                    if "PARTS COMPUTATION" in row_str:
                        score += 50
                        for r2 in range(r + 1, min(r + 15, sheet.max_row + 1)):
                            dim2 = sheet.row_dimensions.get(r2)
                            if dim2 and dim2.hidden:
                                continue
                            r2_vals = [str(sheet.cell(row=r2, column=c).value or "").strip() for c in range(1, min(20, sheet.max_column + 1))]
                            if any("PARTICULAR" in v.upper() for v in r2_vals):
                                h_row = r2
                                for c, v in enumerate(r2_vals):
                                    v_u = v.upper()
                                    if "PARTICULAR" in v_u: cols["particular"] = c + 1
                                    elif "ESTIMATE" in v_u: cols["estimate"] = c + 1
                                    elif "BILL" in v_u: cols["bill_no"] = c + 1
                                    elif "ES." in v_u or v_u == "ES" or "S.NO" in v_u: cols["es"] = c + 1
                                break
                        if h_row > 0:
                            break

                    elif any("PARTICULAR" in v.upper() for v in row_vals) and any("ESTIMATE" in v.upper() for v in row_vals):
                        score += 40
                        h_row = r
                        for c, v in enumerate(row_vals):
                            v_u = v.upper()
                            if "PARTICULAR" in v_u: cols["particular"] = c + 1
                            elif "ESTIMATE" in v_u: cols["estimate"] = c + 1
                            elif "BILL" in v_u: cols["bill_no"] = c + 1
                            elif "ES." in v_u or v_u == "ES" or "S.NO" in v_u: cols["es"] = c + 1
                        break

                if sname.lower() == "sheet1":
                    score += 15
                if target_sheet and sname.strip().lower() == target_sheet.strip().lower():
                    score += 200

                if score > best_score:
                    best_score = score
                    best_sheet = sname
                    best_cols = cols
                    if h_row > 0:
                        best_header_r = h_row
                        best_start_r = h_row + 1
                        for r_data in range(h_row + 1, min(h_row + 15, sheet.max_row + 1)):
                            dim_d = sheet.row_dimensions.get(r_data)
                            if dim_d and dim_d.hidden:
                                continue
                            sub_vals = [str(sheet.cell(row=r_data, column=c).value or "").strip().lower() for c in range(1, min(15, sheet.max_column + 1))]
                            if any(term in sub_vals for term in ["metal", "non metal", "glass", "labour", "tax"]):
                                continue
                            best_start_r = r_data
                            break

            if not best_sheet and sheet_names:
                best_sheet = sheet_names[0]

            # Detect LABOUR COMPUTATION start row in best_sheet
            labour_start_r = 359
            if best_sheet and best_sheet in wb.sheetnames:
                ws_best = wb[best_sheet]
                for r_lab in range(1, min(ws_best.max_row + 1, 600)):
                    r_lab_vals = [str(ws_best.cell(row=r_lab, column=c).value or "").strip().upper() for c in range(1, min(15, ws_best.max_column + 1))]
                    if "LABOUR COMPUTATION" in " ".join(r_lab_vals):
                        for r_under in range(r_lab + 1, min(r_lab + 20, ws_best.max_row + 1)):
                            r_u_vals = [str(ws_best.cell(row=r_under, column=c).value or "").strip().upper() for c in range(1, min(15, ws_best.max_column + 1))]
                            if "ASSESSED AS UNDER" in " ".join(r_u_vals):
                                labour_start_r = r_under + 1
                                break
                            if re.match(r"^1(\.0)?$", r_u_vals[0]):
                                labour_start_r = r_under
                                break
                        break

            res = {
                "success": True,
                "file_path": norm_path,
                "file_type": "xlsx",
                "sheets": sheet_names,
                "selected_sheet": best_sheet,
                "header_row": best_header_r,
                "data_start_row": best_start_r,
                "labour_start_row": labour_start_r,
                "particular_col": best_cols["particular"],
                "estimate_col": best_cols["estimate"],
                "bill_no_col": best_cols["bill_no"],
                "es_col": best_cols["es"],
                "particular_col_letter": openpyxl.utils.get_column_letter(best_cols["particular"]),
                "estimate_col_letter": openpyxl.utils.get_column_letter(best_cols["estimate"])
            }
            _ESTIMATE_INFO_CACHE[cache_key] = (mtime, size, res)
            return res.copy()
        except Exception as e:
            return {"success": False, "error": f"Failed reading Excel file: {str(e)}"}
        finally:
            if wb is not None:
                try:
                    wb.close()
                except Exception:
                    pass


def write_parts_to_excel(file_path, sheet_name=None, parts=None, start_row=None, bill_no=""):
    """
    Writes extracted parts into the PARTS COMPUTATION table:
    - Serial No ('Es.'): Sequential item index (e.g. 1, 2, 3...)
    - Bill No ('Bill No.'): Optional bill reference
    - Particular ('Particular'): Part description (Col C)
    - Estimate ('Estimate'): Taxable Amount (Col D)
    Ensures parts are written strictly to VISIBLE rows, skipping any hidden template rows.
    In .xlsx/.xlsm workbooks, dynamically inserts rows if the number of parts exceeds available table rows.
    """
    if not os.path.exists(file_path):
        return {"success": False, "error": f"File not found: {file_path}"}

    if not parts:
        return {"success": False, "error": "No parts provided to write."}

    # Safety check for file lock
    try:
        with open(file_path, "r+b"):
            pass
    except PermissionError:
        return {
            "success": False,
            "error": "The Excel file is currently open in Microsoft Excel. Please close it first."
        }
    except Exception as e:
        return {"success": False, "error": f"File access error: {str(e)}"}

    # Create safety backup
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    backup_path = f"{file_path}.bak_{timestamp}"
    try:
        shutil.copy2(file_path, backup_path)
        _prune_old_backups(file_path)
    except Exception as e:
        print(f"Warning creating backup: {e}")

    # Inspect sheet and determine exact target columns and start row
    sheet_info = inspect_estimate_excel(file_path, target_sheet=sheet_name)
    if not sheet_info.get("success"):
        return {"success": False, "error": sheet_info.get("error", "Failed to inspect Excel file.")}

    target_sheet_name = sheet_info.get("selected_sheet")
    part_col_1based = sheet_info.get("particular_col", 3)
    est_col_1based = sheet_info.get("estimate_col", 4)
    bill_col_1based = sheet_info.get("bill_no_col", 2)
    es_col_1based = sheet_info.get("es_col", 1)

    detected_visible_start = sheet_info.get("data_start_row", 171)
    if start_row is None or int(start_row) <= 0:
        start_row = detected_visible_start
    else:
        start_row = int(start_row)
        if start_row < detected_visible_start and detected_visible_start >= 100:
            start_row = detected_visible_start

    ext = os.path.splitext(file_path)[1].lower()

    if ext == ".xls":
        try:
            import xlrd
            from xlutils.copy import copy

            rb = xlrd.open_workbook(file_path, formatting_info=True)
            sheet_names = rb.sheet_names()

            sheet_idx = 0
            if target_sheet_name in sheet_names:
                sheet_idx = sheet_names.index(target_sheet_name)
            else:
                for idx, sn in enumerate(sheet_names):
                    if sn.strip().lower() == str(target_sheet_name).strip().lower():
                        sheet_idx = idx
                        target_sheet_name = sn
                        break

            rb_sheet = rb.sheet_by_index(sheet_idx)
            wb = copy(rb)
            sheet = wb.get_sheet(sheet_idx)

            current_r = max(0, start_row - 1)
            c_part = max(0, part_col_1based - 1)
            c_est = max(0, est_col_1based - 1)
            c_bill = max(0, bill_col_1based - 1)
            c_es = max(0, es_col_1based - 1)

            item_num = 1
            written = 0
            first_written_row = None
            last_written_row = None

            for p in parts:
                desc = p.get("particular", "").strip()
                amt = p.get("amount", 0.0)
                try:
                    amt_num = float(amt)
                except (ValueError, TypeError):
                    amt_num = 0.0

                while current_r < rb_sheet.nrows:
                    rinfo = rb_sheet.rowinfo_map.get(current_r) if hasattr(rb_sheet, "rowinfo_map") else None
                    if rinfo and rinfo.hidden:
                        current_r += 1
                    else:
                        break

                if first_written_row is None:
                    first_written_row = current_r + 1
                last_written_row = current_r + 1

                try:
                    sheet.write(current_r, c_es, item_num)
                except Exception:
                    pass

                # User requirement: Leave Column B (Bill No.) blank when pasting from Estimate
                try:
                    sheet.write(current_r, c_bill, "")
                except Exception:
                    pass

                sheet.write(current_r, c_part, desc)
                sheet.write(current_r, c_est, amt_num)

                current_r += 1
                item_num += 1
                written += 1

            wb.save(file_path)
            invalidate_estimate_cache(file_path)

            return {
                "success": True,
                "written_count": written,
                "start_row": first_written_row or start_row,
                "end_row": last_written_row or (start_row + written - 1),
                "backup_created": backup_path,
                "file_path": file_path,
                "sheet": target_sheet_name,
                "particular_col": chr(65 + c_part),
                "estimate_col": chr(65 + c_est)
            }
        except Exception as e:
            return {"success": False, "error": f"Error writing to .xls file: {str(e)}"}

    else:
        # Modern .xlsx / .xlsm using openpyxl
        wb = None
        try:
            import openpyxl

            wb = openpyxl.load_workbook(file_path)
            if target_sheet_name in wb.sheetnames:
                sheet = wb[target_sheet_name]
            else:
                matched = False
                for sn in wb.sheetnames:
                    if sn.strip().lower() == str(target_sheet_name).strip().lower():
                        sheet = wb[sn]
                        target_sheet_name = sn
                        matched = True
                        break
                if not matched:
                    sheet = wb.active
                    target_sheet_name = sheet.title

            current_r = start_row
            c_part = part_col_1based
            c_est = est_col_1based
            c_bill = bill_col_1based
            c_es = es_col_1based

            # Detect if rows exceed existing table space before Total/Labour row
            table_limit_row = None
            for check_r in range(start_row, min(sheet.max_row + 1, start_row + 600)):
                v_particular = str(sheet.cell(row=check_r, column=c_part).value or "").strip().lower()
                v_col1 = str(sheet.cell(row=check_r, column=1).value or "").strip().lower()
                if any(k in v_particular or k in v_col1 for k in ["total", "s.total", "grand total", "labour computation", "total parts"]):
                    table_limit_row = check_r
                    break

            # If parts exceed the available rows before total, dynamically insert rows in openpyxl
            num_parts = len(parts)
            if table_limit_row and (start_row + num_parts >= table_limit_row):
                needed = (start_row + num_parts) - table_limit_row + 2
                sheet.insert_rows(table_limit_row, amount=needed)

            item_num = 1
            written = 0
            first_written_row = None
            last_written_row = None

            for p in parts:
                desc = p.get("particular", "").strip()
                amt = p.get("amount", 0.0)
                try:
                    amt_num = float(amt)
                except (ValueError, TypeError):
                    amt_num = 0.0

                while current_r <= sheet.max_row:
                    dim = sheet.row_dimensions.get(current_r)
                    if dim and dim.hidden:
                        current_r += 1
                    else:
                        break

                if first_written_row is None:
                    first_written_row = current_r
                last_written_row = current_r

                if c_es > 0:
                    sheet.cell(row=current_r, column=c_es, value=item_num)
                # User requirement: Leave Column B (Bill No.) blank when pasting from Estimate
                if c_bill > 0:
                    sheet.cell(row=current_r, column=c_bill, value=None)
                sheet.cell(row=current_r, column=c_part, value=desc)
                sheet.cell(row=current_r, column=c_est, value=amt_num)

                current_r += 1
                item_num += 1
                written += 1

            wb.save(file_path)
            invalidate_estimate_cache(file_path)

            return {
                "success": True,
                "written_count": written,
                "start_row": first_written_row or start_row,
                "end_row": last_written_row or (start_row + written - 1),
                "backup_created": backup_path,
                "file_path": file_path,
                "sheet": target_sheet_name,
                "particular_col": openpyxl.utils.get_column_letter(c_part),
                "estimate_col": openpyxl.utils.get_column_letter(c_est)
            }
        except Exception as e:
            return {"success": False, "error": f"Error writing to Excel: {str(e)}"}
        finally:
            if wb is not None:
                try:
                    wb.close()
                except Exception:
                    pass


def write_labour_to_excel(file_path, sheet_name=None, labour_items=None, start_row=None, bill_no=""):
    """
    Writes extracted labour items into the LABOUR COMPUTATION table:
    - Serial No ('Es.'): Sequential item index (e.g. 1, 2, 3...)
    - Bill No ('Bill No.'): Optional bill reference
    - Particular ('Particular'): Labour description (Col C)
    - Estimate ('Estimate'): Taxable Amount (Col D)
    Ensures items are written strictly to VISIBLE rows, skipping any hidden template rows.
    In .xlsx/.xlsm workbooks, dynamically inserts rows if the number of items exceeds available table rows.
    """
    if not os.path.exists(file_path):
        return {"success": False, "error": f"File not found: {file_path}"}

    if not labour_items:
        return {"success": False, "error": "No labour items provided to write."}

    # Safety check for file lock
    try:
        with open(file_path, "r+b"):
            pass
    except PermissionError:
        return {
            "success": False,
            "error": "The Excel file is currently open in Microsoft Excel. Please close it first."
        }
    except Exception as e:
        return {"success": False, "error": f"File access error: {str(e)}"}

    # Create safety backup
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    backup_path = f"{file_path}.bak_{timestamp}"
    try:
        shutil.copy2(file_path, backup_path)
        _prune_old_backups(file_path)
    except Exception as e:
        print(f"Warning creating backup: {e}")

    # Inspect sheet and determine target columns and start row
    sheet_info = inspect_estimate_excel(file_path, target_sheet=sheet_name)
    if not sheet_info.get("success"):
        return {"success": False, "error": sheet_info.get("error", "Failed to inspect Excel file.")}

    target_sheet_name = sheet_info.get("selected_sheet")
    part_col_1based = sheet_info.get("particular_col", 3)
    est_col_1based = sheet_info.get("estimate_col", 4)
    bill_col_1based = sheet_info.get("bill_no_col", 2)
    es_col_1based = sheet_info.get("es_col", 1)

    detected_start = sheet_info.get("labour_start_row", 359)
    if start_row is None or int(start_row) <= 0:
        start_row = detected_start
    else:
        start_row = int(start_row)

    ext = os.path.splitext(file_path)[1].lower()

    if ext == ".xls":
        try:
            import xlrd
            from xlutils.copy import copy

            rb = xlrd.open_workbook(file_path, formatting_info=True)
            sheet_names = rb.sheet_names()

            sheet_idx = 0
            if target_sheet_name in sheet_names:
                sheet_idx = sheet_names.index(target_sheet_name)
            else:
                for idx, sn in enumerate(sheet_names):
                    if sn.strip().lower() == str(target_sheet_name).strip().lower():
                        sheet_idx = idx
                        target_sheet_name = sn
                        break

            rb_sheet = rb.sheet_by_index(sheet_idx)
            wb = copy(rb)
            sheet = wb.get_sheet(sheet_idx)

            current_r = max(0, start_row - 1)
            c_part = max(0, part_col_1based - 1)
            c_est = max(0, est_col_1based - 1)
            c_bill = max(0, bill_col_1based - 1)
            c_es = max(0, es_col_1based - 1)

            item_num = 1
            written = 0
            first_written_row = None
            last_written_row = None

            for it in labour_items:
                desc = it.get("particular", "").strip()
                amt = it.get("amount", 0.0)
                try:
                    amt_num = float(amt)
                except (ValueError, TypeError):
                    amt_num = 0.0

                while current_r < rb_sheet.nrows:
                    rinfo = rb_sheet.rowinfo_map.get(current_r) if hasattr(rb_sheet, "rowinfo_map") else None
                    if rinfo and rinfo.hidden:
                        current_r += 1
                    else:
                        break

                if first_written_row is None:
                    first_written_row = current_r + 1
                last_written_row = current_r + 1

                try:
                    sheet.write(current_r, c_es, item_num)
                except Exception:
                    pass

                # User requirement: Leave Column B (Bill No.) blank when pasting from Estimate
                try:
                    sheet.write(current_r, c_bill, "")
                except Exception:
                    pass

                sheet.write(current_r, c_part, desc)
                sheet.write(current_r, c_est, amt_num)

                current_r += 1
                item_num += 1
                written += 1

            wb.save(file_path)
            invalidate_estimate_cache(file_path)

            return {
                "success": True,
                "written_count": written,
                "start_row": first_written_row or start_row,
                "end_row": last_written_row or (start_row + written - 1),
                "backup_created": backup_path,
                "file_path": file_path,
                "sheet": target_sheet_name,
                "particular_col": chr(65 + c_part),
                "estimate_col": chr(65 + c_est)
            }
        except Exception as e:
            return {"success": False, "error": f"Error writing labour to .xls file: {str(e)}"}

    else:
        # Modern .xlsx / .xlsm using openpyxl
        wb = None
        try:
            import openpyxl

            wb = openpyxl.load_workbook(file_path)
            if target_sheet_name in wb.sheetnames:
                sheet = wb[target_sheet_name]
            else:
                matched = False
                for sn in wb.sheetnames:
                    if sn.strip().lower() == str(target_sheet_name).strip().lower():
                        sheet = wb[sn]
                        target_sheet_name = sn
                        matched = True
                        break
                if not matched:
                    sheet = wb.active
                    target_sheet_name = sheet.title

            current_r = start_row
            c_part = part_col_1based
            c_est = est_col_1based
            c_bill = bill_col_1based
            c_es = es_col_1based

            table_limit_row = None
            for check_r in range(start_row, min(sheet.max_row + 1, start_row + 200)):
                v_particular = str(sheet.cell(row=check_r, column=c_part).value or "").strip().lower()
                v_col1 = str(sheet.cell(row=check_r, column=1).value or "").strip().lower()
                if any(k in v_particular or k in v_col1 for k in ["total", "s.total", "grand total", "total labour"]):
                    table_limit_row = check_r
                    break

            num_items = len(labour_items)
            if table_limit_row and (start_row + num_items >= table_limit_row):
                needed = (start_row + num_items) - table_limit_row + 2
                sheet.insert_rows(table_limit_row, amount=needed)

            item_num = 1
            written = 0
            first_written_row = None
            last_written_row = None

            for it in labour_items:
                desc = it.get("particular", "").strip()
                amt = it.get("amount", 0.0)
                try:
                    amt_num = float(amt)
                except (ValueError, TypeError):
                    amt_num = 0.0

                while current_r <= sheet.max_row:
                    dim = sheet.row_dimensions.get(current_r)
                    if dim and dim.hidden:
                        current_r += 1
                    else:
                        break

                if first_written_row is None:
                    first_written_row = current_r
                last_written_row = current_r

                if c_es > 0:
                    sheet.cell(row=current_r, column=c_es, value=item_num)
                # User requirement: Leave Column B (Bill No.) blank when pasting from Estimate
                if c_bill > 0:
                    sheet.cell(row=current_r, column=c_bill, value=None)
                sheet.cell(row=current_r, column=c_part, value=desc)
                sheet.cell(row=current_r, column=c_est, value=amt_num)

                current_r += 1
                item_num += 1
                written += 1

            wb.save(file_path)
            invalidate_estimate_cache(file_path)

            return {
                "success": True,
                "written_count": written,
                "start_row": first_written_row or start_row,
                "end_row": last_written_row or (start_row + written - 1),
                "backup_created": backup_path,
                "file_path": file_path,
                "sheet": target_sheet_name,
                "particular_col": openpyxl.utils.get_column_letter(c_part),
                "estimate_col": openpyxl.utils.get_column_letter(c_est)
            }
        except Exception as e:
            return {"success": False, "error": f"Error writing labour to Excel: {str(e)}"}
        finally:
            if wb is not None:
                try:
                    wb.close()
                except Exception:
                    pass



# ==============================================================================
# PART MATERIAL CATEGORIZATION & RECONCILIATION / TALLY WITH FINAL BILL
# ==============================================================================

def categorize_part_material(desc, code=""):
    """
    Categorizes an automotive spare part into:
    - 'Glass' (Col G) - 0% Depr (Windshield, mirrors, lamps, glasses)
    - 'Non Metal' (Col F) - Plastics, Rubber, Fibres, Foam, Consumables
    - 'Metal' (Col E) - Sheet metal panels, Suspension, Engine, Mechanical
    """
    d = (desc or "").upper()
    c = (code or "").upper()
    text = f"{d} {c}"

    # 1. Glass Category
    glass_patterns = [
        r'\bGLASS\b', r'\bWINDSHIELD\b', r'\bWINDSCREEN\b', r'\bREAR\s*SCREEN\b',
        r'\bDOOR\s*GLASS\b', r'\bQUARTER\s*GLASS\b', r'\bVENT\s*GLASS\b',
        r'\bMIRROR\s*GLASS\b', r'\bMIRROR\s*PLATE\b',
        r'\bHEAD\s*LAMP\b', r'\bHEADLAMP\b', r'\bHEADLIGHT\b', r'\bHEAD\s*LIGHT\b',
        r'\bTAIL\s*LAMP\b', r'\bTAILLAMP\b', r'\bTAIL\s*LIGHT\b',
        r'\bFOG\s*LAMP\b', r'\bFOGLAMP\b', r'\bFOG\s*LIGHT\b',
        r'\bROOM\s*LAMP\b', r'\bLAMP\s*ASSY\b', r'\bINDICATOR\b', r'\bBLINKER\b',
        r'\bREFLECTOR\b', r'\bLENS\b'
    ]
    for pat in glass_patterns:
        if re.search(pat, text):
            if re.search(r'\bBRACKET\b|\bCOVER\b|\bHOUSING\b|\bBEZEL\b', text) and not re.search(r'\bUNIT\b|\bASSY\b', text):
                return "Non Metal"
            return "Glass"

    # 2. Non-Metal Category (Plastics, Rubbers, Fibres, Interior, Fluids)
    non_metal_patterns = [
        r'\bBUMPER\b', r'\bCOVER\b', r'\bGRILLE\b', r'\bGRILL\b', r'\bMOULDING\b',
        r'\bMOLDING\b', r'\bGARNISH\b', r'\bCLADDING\b', r'\bFASCIA\b', r'\bLINING\b',
        r'\bMUDGUARD\b', r'\bSPLASH\s*SHIELD\b', r'\bAPRON\b', r'\bARCH\b',
        r'\bDEFLECTOR\b', r'\bSPOILER\b', r'\bDUCT\b', r'\bAIR\s*GUIDE\b',
        r'\bEMBLEM\b', r'\bMONOGRAM\b', r'\bLOGO\b', r'\bBADGE\b', r'\bNAMEPLATE\b',
        r'\bHANDLE\b', r'\bKNOB\b', r'\bTRIM\b', r'\bCONSOLE\b', r'\bDASHBOARD\b',
        r'\bPAD\b', r'\bFOAM\b', r'\bCUSHION\b', r'\bSEAT\b', r'\bCARPET\b', r'\bMAT\b',
        r'\bSUNVISOR\b', r'\bWEATHERSTRIP\b', r'\bRUBBER\b', r'\bBEADING\b',
        r'\bHOSE\b', r'\bPIPE\b', r'\bTUBE\b', r'\bSEAL\b', r'\bGASKET\b', r'\bO-RING\b',
        r'\bBUSH\b', r'\bBOOT\b', r'\bPACKING\b', r'\bMOUNTING\b', r'\bMOUNT\b',
        r'\bCLIP\b', r'\bRETAINER\b', r'\bFASTENER\b', r'\bRIVET\b', r'\bPIN\b',
        r'\bPLUG\b', r'\bCAP\b', r'\bFILTER\b', r'\bELEMENT\b', r'\bOIL\b',
        r'\bCOOLANT\b', r'\bFLUID\b', r'\bPAINT\b', r'\bMATERIAL\b', r'\bADHESIVE\b',
        r'\bSEALANT\b', r'\bPLASTIC\b', r'\bFIBRE\b', r'\bFIBER\b', r'\bTYRE\b',
        r'\bTIRE\b', r'\bBATTERY\b', r'\bWIRING\b', r'\bHARNESS\b', r'\bRELAY\b',
        r'\bSENSOR\b', r'\bHORN\b', r'\bPIECE\b'
    ]
    for pat in non_metal_patterns:
        if re.search(pat, text):
            if re.search(r'\bEXHAUST\s*PIPE\b|\bSTEEL\s*PIPE\b', text):
                return "Metal"
            return "Non Metal"

    # 3. Metal Category (Sheet metal, Suspension, Engine, Steering, Structural)
    metal_patterns = [
        r'\bHOOD\b', r'\bBONNET\b', r'\bFENDER\b', r'\bDOOR\b', r'\bPANEL\b',
        r'\bROOF\b', r'\bPILLAR\b', r'\bQUARTER\s*PANEL\b', r'\bTRUNK\b',
        r'\bBOOT\s*LID\b', r'\bTAILGATE\b', r'\bDICKY\b', r'\bBEAM\b',
        r'\bCROSS\s*MEMBER\b', r'\bSIDE\s*MEMBER\b', r'\bMEMBER\b', r'\bRAIL\b',
        r'\bSILL\b', r'\bRUNNING\s*BOARD\b', r'\bFLOOR\b', r'\bFRAME\b',
        r'\bCHASSIS\b', r'\bSUBFRAME\b', r'\bCRADLE\b', r'\bREINFORCEMENT\b',
        r'\bREINFORCE\b', r'\bSUPPORT\b', r'\bBRACKET\b', r'\bHINGE\b',
        r'\bLOCK\b', r'\bLATCH\b', r'\bSTRIKER\b', r'\bSTAY\b', r'\bSTRUT\b',
        r'\bARM\b', r'\bLINK\b', r'\bROD\b', r'\bTIE\s*ROD\b', r'\bSUSPENSION\b',
        r'\bAXLE\b', r'\bSHAFT\b', r'\bKNUCKLE\b', r'\bHUB\b', r'\bDISC\b',
        r'\bROTOR\b', r'\bDRUM\b', r'\bCALIPER\b', r'\bCYLINDER\b',
        r'\bSTEERING\b', r'\bRACK\b', r'\bCOLUMN\b', r'\bGEARBOX\b',
        r'\bTRANSMISSION\b', r'\bCLUTCH\b', r'\bENGINE\b', r'\bRADIATOR\b',
        r'\bCONDENSER\b', r'\bINTERCOOLER\b', r'\bCOMPRESSOR\b', r'\bEXHAUST\b',
        r'\bSILENCER\b', r'\bMUFFLER\b', r'\bMANIFOLD\b', r'\bSHEET\b',
        r'\bSTEEL\b', r'\bALUMINIUM\b', r'\bALUMINUM\b', r'\bIRON\b', r'\bMETAL\b'
    ]
    for pat in metal_patterns:
        if re.search(pat, text):
            return "Metal"

    return "Metal"


def normalize_part_code(code):
    if not code:
        return ""
    c = str(code).upper().strip()
    c = re.sub(r'^[A-Z]-', '', c)
    return re.sub(r'[^A-Z0-9]', '', c)


def tokenize_desc(desc):
    if not desc:
        return set()
    words = re.findall(r'[A-Za-z0-9]+', desc.lower())
    stop = {'assy', 'sub', 'comp', 'no', 'nos', 'rh', 'lh', 'fr', 'rr', 'front', 'rear', 'right', 'left', 'upper', 'lower', 'lwr', 'upr'}
    meaningful = [w for w in words if w not in stop and len(w) > 1]
    return set(meaningful) if meaningful else set(words)


def tally_estimate_with_bill(estimate_parts, bill_parts):
    """
    Tallies each estimate part against final bill parts.
    - If in both Estimate and Final Bill:
      Categorizes as Metal (Col E), Non Metal (Col F), or Glass (Col G).
      Cost = Final bill taxable amount (or estimate amount if bill is 0).
    - If in Estimate but NOT in Final Bill:
      Categorizes as 'Not Allowed' (Yellow Col H).
      Cost = Estimate taxable amount.
    """
    bill_available = list(bill_parts or [])
    used_bill_indices = set()
    tallied = []

    for idx, est in enumerate(estimate_parts or []):
        est_desc = est.get("particular", "")
        est_code = est.get("code", "")
        est_amt = float(est.get("amount", 0.0) or 0.0)
        est_code_norm = normalize_part_code(est_code)
        est_tokens = tokenize_desc(est_desc)

        best_match_idx = None
        best_match_score = 0.0

        # Pass 1: Exact / Normalized Code Match
        if est_code_norm and len(est_code_norm) >= 4:
            for b_idx, b in enumerate(bill_available):
                if b_idx in used_bill_indices:
                    continue
                b_code_norm = normalize_part_code(b.get("code", ""))
                if b_code_norm and (b_code_norm == est_code_norm or b_code_norm in est_code_norm or est_code_norm in b_code_norm):
                    best_match_idx = b_idx
                    best_match_score = 1.0
                    break

        # Pass 2: Description token overlap if code did not match
        if best_match_idx is None and est_tokens:
            for b_idx, b in enumerate(bill_available):
                if b_idx in used_bill_indices:
                    continue
                b_tokens = tokenize_desc(b.get("particular", ""))
                if not b_tokens:
                    continue
                inter = est_tokens & b_tokens
                if inter:
                    score = len(inter) / max(len(est_tokens), len(b_tokens))
                    if score > best_match_score and (score >= 0.50 or len(inter) >= 2):
                        best_match_score = score
                        best_match_idx = b_idx

        # Parse estimate quantity
        try:
            est_qty = float(est.get("qty", 1.0) or 1.0)
            if est_qty <= 0:
                est_qty = 1.0
        except (ValueError, TypeError):
            est_qty = 1.0

        if best_match_idx is not None:
            used_bill_indices.add(best_match_idx)
            matched_b = bill_available[best_match_idx]
            b_amt = float(matched_b.get("amount", 0.0) or 0.0)

            # Parse bill quantity
            try:
                b_qty = float(matched_b.get("qty", 1.0) or 1.0)
                if b_qty <= 0:
                    b_qty = 1.0
            except (ValueError, TypeError):
                b_qty = 1.0

            # Calculate unit price from bill
            b_unit = float(matched_b.get("unit_price", 0.0) or 0.0)
            if b_unit <= 0 and b_qty > 0 and b_amt > 0:
                b_unit = b_amt / b_qty
            elif b_unit <= 0 and est_qty > 0 and est_amt > 0:
                b_unit = est_amt / est_qty
            elif b_unit <= 0:
                b_unit = b_amt

            # Bill serial number tracking
            raw_srl = matched_b.get("srl") or matched_b.get("sno") or (best_match_idx + 1)
            b_srl = re.sub(r'[\.\s]', '', str(raw_srl)) or str(best_match_idx + 1)

            # Limit bill quantity to estimate quantity (estimate is the upper bound)
            allowed_qty = min(b_qty, est_qty)
            allowed_cost = round(allowed_qty * b_unit, 2) if b_unit > 0 else (b_amt if b_amt > 0 else est_amt)

            # Excess quantity goes to Column H (Not Allowed)
            excess_qty = max(0.0, round(b_qty - allowed_qty, 2))
            excess_cost = round(excess_qty * b_unit, 2)

            cat = categorize_part_material(est_desc, est_code)
            target_col = "E" if cat == "Metal" else ("F" if cat == "Non Metal" else "G")

            tallied.append({
                "index": idx + 1,
                "est_particular": est_desc,
                "est_desc": est_desc,
                "est_code": est_code,
                "est_qty": est_qty,
                "est_amount": est_amt,
                "matched": True,
                "bill_srl": b_srl,
                "bill_particular": matched_b.get("particular", ""),
                "bill_desc": matched_b.get("particular", ""),
                "bill_code": matched_b.get("code", ""),
                "bill_qty": b_qty,
                "bill_amount": b_amt,
                "unit_price": round(b_unit, 2),
                "qty": allowed_qty,
                "allowed_qty": allowed_qty,
                "excess_qty": excess_qty,
                "max_qty": est_qty,
                "category": cat,
                "target_col": target_col,
                "final_cost": round(allowed_cost, 2),
                "allowed_cost": round(allowed_cost, 2),
                "excess_cost": round(excess_cost, 2)
            })
        else:
            # Not in Final Bill -> Disallowed / Yellow Col H
            est_unit = float(est.get("unit_price", 0.0) or 0.0)
            if est_unit <= 0 and est_qty > 0 and est_amt > 0:
                est_unit = est_amt / est_qty
            elif est_unit <= 0:
                est_unit = est_amt

            tallied.append({
                "index": idx + 1,
                "est_particular": est_desc,
                "est_desc": est_desc,
                "est_code": est_code,
                "est_qty": est_qty,
                "est_amount": est_amt,
                "matched": False,
                "bill_srl": "",
                "bill_particular": "",
                "bill_desc": "",
                "bill_code": "",
                "bill_qty": 0.0,
                "bill_amount": 0.0,
                "unit_price": round(est_unit, 2),
                "qty": est_qty,
                "allowed_qty": 0.0,
                "excess_qty": 0.0,
                "max_qty": est_qty,
                "category": "Not Allowed",
                "target_col": "H",
                "final_cost": round(est_amt, 2),
                "allowed_cost": 0.0,
                "excess_cost": 0.0
            })

    return tallied


def write_tallied_costs_to_excel(file_path, sheet_name=None, tallied_items=None, start_row=None):
    """
    Writes reconciled / tallied costs into the PARTS COMPUTATION table:
    - Col E (Metal): if category == 'Metal'
    - Col F (Non Metal): if category == 'Non Metal'
    - Col G (Glass): if category == 'Glass'
    - Col H (Yellow / Not Allowed): if category == 'Not Allowed'
    Leaves Col D (Estimate) and Col B (Bill No) untouched.
    """
    if not file_path or not os.path.exists(file_path):
        return {"success": False, "error": f"Target Assessment Excel not found: {file_path}"}

    if not tallied_items:
        return {"success": False, "error": "No tallied items to write."}

    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    backup_path = f"{file_path}.bak_{timestamp}"
    try:
        shutil.copy2(file_path, backup_path)
        _prune_old_backups(file_path)
    except Exception as e:
        print(f"Warning creating backup: {e}")

    ext = os.path.splitext(file_path)[1].lower()
    is_xls = ext == ".xls"

    sheet_info = inspect_estimate_excel(file_path, target_sheet=sheet_name)
    if not sheet_info.get("success"):
        return {"success": False, "error": sheet_info.get("error", "Failed to inspect sheet.")}

    target_sheet_name = sheet_info.get("selected_sheet")
    detected_visible_start = sheet_info.get("data_start_row", 171)
    if start_row is None or int(start_row) <= 0:
        start_row = detected_visible_start
    else:
        start_row = int(start_row)

    col_map = {
        "Metal": 4,       # Col E (0-based in xlwt)
        "Non Metal": 5,   # Col F
        "Glass": 6,       # Col G
        "Not Allowed": 7  # Col H
    }

    metal_count = sum(1 for it in tallied_items if it.get("category") == "Metal")
    non_metal_count = sum(1 for it in tallied_items if it.get("category") == "Non Metal")
    glass_count = sum(1 for it in tallied_items if it.get("category") == "Glass")
    disallowed_count = sum(1 for it in tallied_items if it.get("category") == "Not Allowed")

    if is_xls:
        import xlrd
        from xlutils.copy import copy
        rb = xlrd.open_workbook(file_path, formatting_info=True)
        wb = copy(rb)
        sheet_idx = 0
        if target_sheet_name:
            for i, sname in enumerate(rb.sheet_names()):
                if sname.lower() == target_sheet_name.lower():
                    sheet_idx = i
                    break
        sheet = wb.get_sheet(sheet_idx)
        rb_sheet = rb.sheet_by_index(sheet_idx)

        current_r = start_row - 1
        written = 0
        first_r = current_r + 1
        last_r = current_r

        for it in tallied_items:
            cat = it.get("category", "Not Allowed")
            cost = float(it.get("final_cost", 0.0) or it.get("cost", 0.0) or 0.0)
            excess_cost = float(it.get("excess_cost", 0.0) or 0.0)
            target_c = col_map.get(cat, 7)

            part_name = str(it.get("est_particular") or it.get("est_desc") or "").strip()
            est_amt = float(it.get("est_amount", 0.0) or 0.0)

            # Col B: Bill serial number for confirmation (or blank if unmatched)
            b_srl = it.get("bill_srl")
            if it.get("matched") and b_srl:
                try:
                    srl_val = int(b_srl)
                except (ValueError, TypeError):
                    srl_val = str(b_srl)
                sheet.write(current_r, 1, srl_val)
            else:
                sheet.write(current_r, 1, "")

            # Ensure Col C has the part name
            c_val = rb_sheet.cell_value(current_r, 2) if current_r < rb_sheet.nrows else ""
            if not str(c_val).strip() and part_name:
                sheet.write(current_r, 2, part_name)

            # Ensure Col D has estimate amount
            d_val = rb_sheet.cell_value(current_r, 3) if current_r < rb_sheet.nrows else ""
            if (d_val == "" or d_val is None) and est_amt > 0:
                sheet.write(current_r, 3, est_amt)

            # Write costs to category columns:
            # If matched and allowed: Col E/F/G gets allowed cost, and Col H gets excess_cost if any
            # If unmatched: Col E/F/G are blank, Col H gets full cost / estimate amount
            if it.get("matched") and cat != "Not Allowed":
                for c_idx in [4, 5, 6]:
                    if c_idx == target_c:
                        sheet.write(current_r, c_idx, cost)
                    else:
                        sheet.write(current_r, c_idx, "")
                if excess_cost > 0:
                    sheet.write(current_r, 7, excess_cost)
                else:
                    sheet.write(current_r, 7, "")
            else:
                for c_idx in [4, 5, 6]:
                    sheet.write(current_r, c_idx, "")
                sheet.write(current_r, 7, cost if cost > 0 else est_amt)

            last_r = current_r + 1
            current_r += 1
            written += 1

        wb.save(file_path)
        invalidate_estimate_cache(file_path)
        return {
            "success": True,
            "written_count": written,
            "start_row": first_r,
            "end_row": last_r,
            "sheet": target_sheet_name,
            "backup_created": backup_path,
            "metal_count": metal_count,
            "non_metal_count": non_metal_count,
            "glass_count": glass_count,
            "disallowed_count": disallowed_count
        }
    else:
        import openpyxl
        wb = openpyxl.load_workbook(file_path)
        sheet = wb[target_sheet_name] if target_sheet_name in wb.sheetnames else wb.active

        col_map_openpyxl = {
            "Metal": 5,       # Col E (1-based in openpyxl)
            "Non Metal": 6,   # Col F
            "Glass": 7,       # Col G
            "Not Allowed": 8  # Col H
        }

        current_r = start_row
        written = 0
        first_r = current_r
        last_r = current_r

        for it in tallied_items:
            cat = it.get("category", "Not Allowed")
            cost = float(it.get("final_cost", 0.0) or it.get("cost", 0.0) or 0.0)
            excess_cost = float(it.get("excess_cost", 0.0) or 0.0)
            target_col = col_map_openpyxl.get(cat, 8)

            part_name = str(it.get("est_particular") or it.get("est_desc") or "").strip()
            est_amt = float(it.get("est_amount", 0.0) or 0.0)

            # Col B: Bill serial number for confirmation (or blank if unmatched)
            b_srl = it.get("bill_srl")
            if it.get("matched") and b_srl:
                try:
                    srl_val = int(b_srl)
                except (ValueError, TypeError):
                    srl_val = str(b_srl)
                sheet.cell(row=current_r, column=2, value=srl_val)
            else:
                sheet.cell(row=current_r, column=2, value=None)

            # Ensure Col C has the part name
            c_cell = sheet.cell(row=current_r, column=3)
            if not str(c_cell.value or "").strip() and part_name:
                c_cell.value = part_name

            # Ensure Col D has estimate amount
            d_cell = sheet.cell(row=current_r, column=4)
            if (d_cell.value is None or d_cell.value == "") and est_amt > 0:
                d_cell.value = est_amt

            # Write costs
            if it.get("matched") and cat != "Not Allowed":
                for c_idx in [5, 6, 7]:
                    if c_idx == target_col:
                        sheet.cell(row=current_r, column=c_idx, value=cost)
                    else:
                        sheet.cell(row=current_r, column=c_idx, value=None)
                if excess_cost > 0:
                    sheet.cell(row=current_r, column=8, value=excess_cost)
                else:
                    sheet.cell(row=current_r, column=8, value=None)
            else:
                for c_idx in [5, 6, 7]:
                    sheet.cell(row=current_r, column=c_idx, value=None)
                sheet.cell(row=current_r, column=8, value=cost if cost > 0 else est_amt)

            last_r = current_r
            current_r += 1
            written += 1

        wb.save(file_path)
        invalidate_estimate_cache(file_path)
        return {
            "success": True,
            "written_count": written,
            "start_row": first_r,
            "end_row": last_r,
            "sheet": target_sheet_name,
            "backup_created": backup_path,
            "metal_count": metal_count,
            "non_metal_count": non_metal_count,
            "glass_count": glass_count,
            "disallowed_count": disallowed_count
        }



# ==============================================================================
# LABOUR CATEGORIZATION & RECONCILIATION / TALLY WITH FINAL BILL
# ==============================================================================

def categorize_labour(desc, code=""):
    """
    Categorizes an automotive labour charge into:
    - 'Painting' (Col G) - Paint preparation, refinishing, top coat, buffing, polishing, rubbing
    - 'Denting' (Col F) - Sheet metal dent repair, panel beating, straightening, pulling, welding
    - 'Opening/rft' (Col E) - R&R, removal, refitting, overhaul, dismantling, replacement, testing, scan, alignment
    """
    d = (desc or "").upper()
    c = (code or "").upper()
    text = f"{d} {c}"

    # 1. Painting Category (Col G)
    paint_patterns = [
        r'\bPAINT\b', r'\bPAINTING\b', r'\bREPAINT\b', r'\bREFINISH\b',
        r'\bSURFACER\b', r'\bBUFFING\b', r'\bPOLISHING\b', r'\bRUBBING\b',
        r'\bCLEAR\s*COAT\b', r'\bTOUCH\s*UP\b', r'\bPRIMER\b', r'\bCOLOR\b',
        r'\bBAKE\b', r'\bBOOTH\b', r'\bTOP\s*COAT\b'
    ]
    for pat in paint_patterns:
        if re.search(pat, text):
            return "Painting"

    # 2. Denting Category (Col F)
    dent_patterns = [
        r'\bDENT\b', r'\bDENTING\b', r'\bPANEL\s*BEATING\b', r'\bSTRAIGHTENING\b',
        r'\bPULLING\b', r'\bTINKERING\b', r'\bWELDING\b', r'\bBENCH\b',
        r'\bCHASSIS\s*ALIGNMENT\b', r'\bFRAME\s*ALIGNMENT\b'
    ]
    for pat in dent_patterns:
        if re.search(pat, text):
            return "Denting"

    # 3. Opening/rft Category (Col E) - Default for R&R, removal, replacement, overhaul, mechanical, alignment, or general labour
    return "Opening/rft"


def tally_estimate_labour_with_bill(estimate_labour, bill_labour):
    """
    Tallies each estimate labour item against final bill labour items.
    - If in both Estimate and Final Bill:
      Categorizes as 'Opening/rft' (Col E), 'Denting' (Col F), or 'Painting' (Col G).
      Quantity limited to min(bill_qty, est_qty).
      Cost = Final bill labour amount (or allowed_qty * bill_unit_price).
    - If in Estimate but NOT in Final Bill:
      Categorizes as 'Not Allowed' (Yellow Col H).
      Cost = Estimate labour taxable amount.
    """
    bill_available = list(bill_labour or [])
    used_bill_indices = set()
    tallied = []

    for idx, est in enumerate(estimate_labour or []):
        est_desc = est.get("particular", "")
        est_code = est.get("code", "")
        est_amt = float(est.get("amount", 0.0) or 0.0)
        est_code_norm = normalize_part_code(est_code)
        est_tokens = tokenize_desc(est_desc)

        best_match_idx = None
        best_match_score = 0.0

        # Pass 1: Code Match (if code is meaningful)
        if est_code_norm and len(est_code_norm) >= 3:
            for b_idx, b in enumerate(bill_available):
                if b_idx in used_bill_indices:
                    continue
                b_code_norm = normalize_part_code(b.get("code", ""))
                if b_code_norm and (b_code_norm == est_code_norm or b_code_norm in est_code_norm or est_code_norm in b_code_norm):
                    best_match_idx = b_idx
                    best_match_score = 1.0
                    break

        # Pass 2: Description token overlap if code did not match
        if best_match_idx is None and est_tokens:
            for b_idx, b in enumerate(bill_available):
                if b_idx in used_bill_indices:
                    continue
                b_tokens = tokenize_desc(b.get("particular", ""))
                if not b_tokens:
                    continue
                inter = est_tokens & b_tokens
                if inter:
                    score = len(inter) / max(len(est_tokens), len(b_tokens))
                    if score > best_match_score and (score >= 0.40 or len(inter) >= 1):
                        best_match_score = score
                        best_match_idx = b_idx

        # Parse estimate quantity
        try:
            est_qty = float(est.get("qty", 1.0) or 1.0)
            if est_qty <= 0:
                est_qty = 1.0
        except (ValueError, TypeError):
            est_qty = 1.0

        if best_match_idx is not None:
            used_bill_indices.add(best_match_idx)
            matched_b = bill_available[best_match_idx]
            b_amt = float(matched_b.get("amount", 0.0) or 0.0)

            # Bill serial number tracking
            raw_srl = matched_b.get("srl") or matched_b.get("sno") or (best_match_idx + 1)
            b_srl = re.sub(r'[\.\s]', '', str(raw_srl)) or str(best_match_idx + 1)

            try:
                b_qty = float(matched_b.get("qty", 1.0) or 1.0)
                if b_qty <= 0:
                    b_qty = 1.0
            except (ValueError, TypeError):
                b_qty = 1.0

            b_unit = float(matched_b.get("unit_price", 0.0) or 0.0)
            if b_unit <= 0 and b_qty > 0 and b_amt > 0:
                b_unit = b_amt / b_qty
            elif b_unit <= 0 and est_qty > 0 and est_amt > 0:
                b_unit = est_amt / est_qty
            elif b_unit <= 0:
                b_unit = b_amt

            allowed_qty = min(b_qty, est_qty)
            allowed_cost = round(allowed_qty * b_unit, 2) if b_unit > 0 else (b_amt if b_amt > 0 else est_amt)
            excess_qty = max(0.0, round(b_qty - allowed_qty, 2))
            excess_cost = round(excess_qty * b_unit, 2)

            cat = categorize_labour(est_desc, est_code)
            target_col = "E" if cat == "Opening/rft" else ("F" if cat == "Denting" else "G")

            tallied.append({
                "index": idx + 1,
                "est_particular": est_desc,
                "est_desc": est_desc,
                "est_code": est_code,
                "est_qty": est_qty,
                "est_amount": est_amt,
                "matched": True,
                "bill_srl": b_srl,
                "bill_particular": matched_b.get("particular", ""),
                "bill_desc": matched_b.get("particular", ""),
                "bill_code": matched_b.get("code", ""),
                "bill_qty": b_qty,
                "bill_amount": b_amt,
                "unit_price": round(b_unit, 2),
                "qty": allowed_qty,
                "allowed_qty": allowed_qty,
                "excess_qty": excess_qty,
                "max_qty": est_qty,
                "category": cat,
                "target_col": target_col,
                "final_cost": round(allowed_cost, 2),
                "allowed_cost": round(allowed_cost, 2),
                "excess_cost": round(excess_cost, 2),
                "item_type": "Labour"
            })
        else:
            # Not in Final Bill -> Disallowed / Col H
            est_unit = float(est.get("unit_price", 0.0) or 0.0)
            if est_unit <= 0 and est_qty > 0 and est_amt > 0:
                est_unit = est_amt / est_qty
            elif est_unit <= 0:
                est_unit = est_amt

            tallied.append({
                "index": idx + 1,
                "est_particular": est_desc,
                "est_desc": est_desc,
                "est_code": est_code,
                "est_qty": est_qty,
                "est_amount": est_amt,
                "matched": False,
                "bill_srl": "",
                "bill_particular": "",
                "bill_desc": "",
                "bill_code": "",
                "bill_qty": 0.0,
                "bill_amount": 0.0,
                "unit_price": round(est_unit, 2),
                "qty": est_qty,
                "allowed_qty": 0.0,
                "excess_qty": 0.0,
                "max_qty": est_qty,
                "category": "Not Allowed",
                "target_col": "H",
                "final_cost": round(est_amt, 2),
                "allowed_cost": 0.0,
                "excess_cost": 0.0,
                "item_type": "Labour"
            })

    return tallied


def write_tallied_labour_to_excel(file_path, sheet_name=None, tallied_labour=None, start_row=None):
    """
    Writes reconciled / tallied labour costs into the LABOUR COMPUTATION table:
    - Col E (Opening/rft): if category == 'Opening/rft'
    - Col F (Denting): if category == 'Denting'
    - Col G (Painting): if category == 'Painting'
    - Col H (Not Allowed): if category == 'Not Allowed'
    Leaves Col D (Estimate) and Col B (Bill No) untouched.
    """
    if not file_path or not os.path.exists(file_path):
        return {"success": False, "error": f"Target Assessment Excel not found: {file_path}"}

    if not tallied_labour:
        return {"success": False, "error": "No tallied labour items to write."}

    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    backup_path = f"{file_path}.bak_{timestamp}"
    try:
        shutil.copy2(file_path, backup_path)
        _prune_old_backups(file_path)
    except Exception as e:
        print(f"Warning creating backup: {e}")

    ext = os.path.splitext(file_path)[1].lower()
    is_xls = ext == ".xls"

    sheet_info = inspect_estimate_excel(file_path, target_sheet=sheet_name)
    if not sheet_info.get("success"):
        return {"success": False, "error": sheet_info.get("error", "Failed to inspect sheet.")}

    target_sheet_name = sheet_info.get("selected_sheet")
    detected_labour_start = sheet_info.get("labour_start_row", 359)
    if start_row is None or int(start_row) <= 0:
        start_row = detected_labour_start
    else:
        start_row = int(start_row)

    col_map = {
        "Opening/rft": 4,  # Col E (0-based)
        "Denting": 5,      # Col F
        "Painting": 6,     # Col G
        "Not Allowed": 7   # Col H
    }

    opening_count = sum(1 for it in tallied_labour if it.get("category") == "Opening/rft")
    denting_count = sum(1 for it in tallied_labour if it.get("category") == "Denting")
    painting_count = sum(1 for it in tallied_labour if it.get("category") == "Painting")
    disallowed_count = sum(1 for it in tallied_labour if it.get("category") == "Not Allowed")

    if is_xls:
        import xlrd
        from xlutils.copy import copy
        rb = xlrd.open_workbook(file_path, formatting_info=True)
        wb = copy(rb)
        sheet_idx = 0
        if target_sheet_name:
            for i, sname in enumerate(rb.sheet_names()):
                if sname.lower() == target_sheet_name.lower():
                    sheet_idx = i
                    break
        sheet = wb.get_sheet(sheet_idx)
        rb_sheet = rb.sheet_by_index(sheet_idx)

        current_r = start_row - 1
        written = 0
        first_r = current_r + 1
        last_r = current_r

        for it in tallied_labour:
            cat = it.get("category", "Not Allowed")
            cost = float(it.get("final_cost", 0.0) or it.get("cost", 0.0) or 0.0)
            excess_cost = float(it.get("excess_cost", 0.0) or 0.0)
            target_c = col_map.get(cat, 7)

            lab_name = str(it.get("est_particular") or it.get("est_desc") or "").strip()
            est_amt = float(it.get("est_amount", 0.0) or 0.0)

            # Col B: Bill serial number for confirmation (or blank if unmatched)
            b_srl = it.get("bill_srl")
            if it.get("matched") and b_srl:
                try:
                    srl_val = int(b_srl)
                except (ValueError, TypeError):
                    srl_val = str(b_srl)
                sheet.write(current_r, 1, srl_val)
            else:
                sheet.write(current_r, 1, "")

            # Ensure Col C has the labour description
            c_val = rb_sheet.cell_value(current_r, 2) if current_r < rb_sheet.nrows else ""
            if not str(c_val).strip() and lab_name:
                sheet.write(current_r, 2, lab_name)

            # Ensure Col D has estimate amount
            d_val = rb_sheet.cell_value(current_r, 3) if current_r < rb_sheet.nrows else ""
            if (d_val == "" or d_val is None) and est_amt > 0:
                sheet.write(current_r, 3, est_amt)

            # Write costs to category columns
            if it.get("matched") and cat != "Not Allowed":
                for c_idx in [4, 5, 6]:
                    if c_idx == target_c:
                        sheet.write(current_r, c_idx, cost)
                    else:
                        sheet.write(current_r, c_idx, "")
                if excess_cost > 0:
                    sheet.write(current_r, 7, excess_cost)
                else:
                    sheet.write(current_r, 7, "")
            else:
                for c_idx in [4, 5, 6]:
                    sheet.write(current_r, c_idx, "")
                sheet.write(current_r, 7, cost if cost > 0 else est_amt)

            last_r = current_r + 1
            current_r += 1
            written += 1

        wb.save(file_path)
        invalidate_estimate_cache(file_path)
        return {
            "success": True,
            "written_count": written,
            "start_row": first_r,
            "end_row": last_r,
            "sheet": target_sheet_name,
            "backup_created": backup_path,
            "opening_count": opening_count,
            "denting_count": denting_count,
            "painting_count": painting_count,
            "disallowed_count": disallowed_count
        }
    else:
        import openpyxl
        wb = openpyxl.load_workbook(file_path)
        sheet = wb[target_sheet_name] if target_sheet_name in wb.sheetnames else wb.active

        col_map_openpyxl = {
            "Opening/rft": 5,  # Col E (1-based)
            "Denting": 6,      # Col F
            "Painting": 7,     # Col G
            "Not Allowed": 8   # Col H
        }

        current_r = start_row
        written = 0
        first_r = current_r
        last_r = current_r

        for it in tallied_labour:
            cat = it.get("category", "Not Allowed")
            cost = float(it.get("final_cost", 0.0) or it.get("cost", 0.0) or 0.0)
            excess_cost = float(it.get("excess_cost", 0.0) or 0.0)
            target_col = col_map_openpyxl.get(cat, 8)

            lab_name = str(it.get("est_particular") or it.get("est_desc") or "").strip()
            est_amt = float(it.get("est_amount", 0.0) or 0.0)

            # Col B: Bill serial number for confirmation (or blank if unmatched)
            b_srl = it.get("bill_srl")
            if it.get("matched") and b_srl:
                try:
                    srl_val = int(b_srl)
                except (ValueError, TypeError):
                    srl_val = str(b_srl)
                sheet.cell(row=current_r, column=2, value=srl_val)
            else:
                sheet.cell(row=current_r, column=2, value=None)

            # Ensure Col C has the labour description
            c_cell = sheet.cell(row=current_r, column=3)
            if not str(c_cell.value or "").strip() and lab_name:
                c_cell.value = lab_name

            # Ensure Col D has estimate amount
            d_cell = sheet.cell(row=current_r, column=4)
            if (d_cell.value is None or d_cell.value == "") and est_amt > 0:
                d_cell.value = est_amt

            # Write costs
            if it.get("matched") and cat != "Not Allowed":
                for c_idx in [5, 6, 7]:
                    if c_idx == target_col:
                        sheet.cell(row=current_r, column=c_idx, value=cost)
                    else:
                        sheet.cell(row=current_r, column=c_idx, value=None)
                if excess_cost > 0:
                    sheet.cell(row=current_r, column=8, value=excess_cost)
                else:
                    sheet.cell(row=current_r, column=8, value=None)
            else:
                for c_idx in [5, 6, 7]:
                    sheet.cell(row=current_r, column=c_idx, value=None)
                sheet.cell(row=current_r, column=8, value=cost if cost > 0 else est_amt)

            last_r = current_r
            current_r += 1
            written += 1

        wb.save(file_path)
        invalidate_estimate_cache(file_path)
        return {
            "success": True,
            "written_count": written,
            "start_row": first_r,
            "end_row": last_r,
            "sheet": target_sheet_name,
            "backup_created": backup_path,
            "opening_count": opening_count,
            "denting_count": denting_count,
            "painting_count": painting_count,
            "disallowed_count": disallowed_count
        }
