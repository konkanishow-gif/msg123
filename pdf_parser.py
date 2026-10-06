import os
import json
import re
import pdfplumber
from pypdf import PdfReader

def load_master_suppliers():
    json_path = os.path.join(os.path.dirname(__file__), 'static', 'suppliers.json')
    if os.path.exists(json_path):
        try:
            with open(json_path, 'r', encoding='utf-8') as f:
                return json.load(f)
        except Exception:
            pass
    return []

MASTER_SUPPLIERS = load_master_suppliers()

COMMON_ITEM_WORDS = {
    "THREADOLET", "SPECTACLE", "UNION", "ELBOW", "FLANGE", "GASKET",
    "NIPPOLET", "WELDOLET", "SOCKOLET", "VALVE", "REDUCER", "COUPLING",
    "NIPPLE", "PIPE", "STUD", "BOLT", "NUT", "BLEED", "ORIFICE",
    "INSULATING", "SWAGELOK", "NEEDLE", "TEE", "CAP", "CROSS", "BUSHING", "PLUG"
}

REMARK_PATTERN = re.compile(
    r"\b(machined(?:\s+from)?|modified(?:\s+from)?|offered|made\s+from|fabricated(?:\s+from)?|locally)\b.*",
    re.IGNORECASE
)

MAKE_MAPPINGS = [
    (r"OMSA[-_]?STOCK|OMSA", "OMSA GERMANY"),
    (r"JD[-_]?STOCK|\bDELCORTE\b|MEGA\/JD", "DELCORTE FRANCE"),
    (r"W[-_]?STOCK|WMA?ASS", "WMASS GERMANY"),
    (r"TK[-_]?STOCK", "TK KOREA"),
    (r"ST&H(?:[-_]?STOCK)?|STH", "ST&H KOREA"),
    (r"\bMFF\b|METALFAR", "Metalfar Italy"),
    (r"BENKAN", "BENKAN THAILAND"),
    (r"ULMA", "ULMA SPAIN"),
    (r"MELESI", "MELESI ITALY"),
    (r"\bMEGA\b", "MEGA ITALY"),
]

def auto_detect_make(raw_text):
    """
    Auto-detects manufacturer / MAKE name from raw item description or product code.
    Mappings:
      - MFF / METALFAR -> Metalfar Italy
      - BENKAN -> BENKAN THAILAND
      - ST&H / STH -> ST&H KOREA
      - W-STOCK / WMAASS / WMASS -> WMASS GERMANY
      - JD-STOCK / DELCORTE / MEGA/JD -> DELCORTE FRANCE
      - OMSA-STOCK / OMSA -> OMSA GERMANY
      - TK-STOCK -> TK KOREA
      - ULMA -> ULMA SPAIN
      - MELESI -> MELESI ITALY
      - MEGA -> MEGA ITALY
    """
    if not raw_text:
        return ""
    for pattern, make_val in MAKE_MAPPINGS:
        if re.search(pattern, raw_text, re.IGNORECASE):
            return make_val
    return ""

def strip_leading_item_number(description, sl_no=None):
    """
    Strips leading item sequence numbers (e.g. '1 ', '2 ', '1. ', '01 ') from the beginning of descriptions,
    while preserving actual size measurements like '1"', '1/2"', '2"'.
    """
    if not description:
        return ""
    desc = description.strip()
    if sl_no is not None:
        pattern = rf"^0*{sl_no}\s*[\.\-]?\s+(?![\"\/])"
        desc = re.sub(pattern, "", desc, flags=re.IGNORECASE).strip()
    # Also strip generic leading index number if followed by dot/dash
    desc = re.sub(r"^\d{1,3}\s*[\.\-]\s+(?![\"\/])", "", desc).strip()
    return desc

def strip_internal_product_code(description):
    """
    Strips internal product codes like LBL25300RF, LTO156KA350LF2, LUN256KNPTA105N,
    XEL90250S40SRWPB, BL151500RF, EL90256KNPTA105JD-STOCK from the beginning of item descriptions.
    Preserves standard fitting names (THREADOLET, SPECTACLE, UNION, ELBOW, etc.).
    """
    if not description:
        return ""
    desc = description.strip()
    tokens = desc.split(maxsplit=1)
    if tokens and len(tokens) > 1:
        first_token = tokens[0].strip()
        if first_token.upper() not in COMMON_ITEM_WORDS:
            # Alphanumeric codes starting with 2+ uppercase letters, length 5-45 (supporting &, -, _, /)
            if re.match(r"^[A-Z]{2,}[A-Z0-9\-_\/&]{3,45}$", first_token):
                desc = tokens[1].strip()
    return desc

def clean_cell_text(v):
    if not v:
        return ""
    return str(v).replace("\n", " ").strip()

def _extract_items_from_tables(pdf):
    """
    Extracts line items using pdfplumber structured tables across all pages.
    Guarantees column isolation (#, description, client_po_item_no, qty, rate, amount)
    and seamlessly merges multi-page split descriptions.
    """
    rows = []
    for p in pdf.pages:
        tables = p.extract_tables()
        for t in tables:
            for r in t:
                cr = [clean_cell_text(c) for c in r]
                if any(cr):
                    rows.append(cr)
                    
    items = []
    current_item = None
    
    for r in rows:
        row_str = " ".join(r).lower()
        if "item & description" in row_str or ("client po" in row_str and "rate" in row_str):
            continue
        if re.search(r"\b(sub total|zero rate|total|standard rate)\b", row_str) and not re.search(r"\b(valve|gasket|bolt|flange|elbow|pipe|fitting)\b", row_str):
            continue
        if re.search(r"\b(duty|vat|tax)\s*\d+%", row_str):
            continue
            
        non_empty = [c for c in r if c]
        if not non_empty:
            continue
            
        # Ignore standalone page number rows e.g. ['2'] or ['3', '', '']
        if len(non_empty) == 1 and re.match(r"^\d{1,3}$", non_empty[0]):
            continue
            
        # Check if first non-empty cell is an item sequence number (1, 2, ..., 24, etc.)
        first_col_num = None
        sl_col_idx = -1
        for idx, col_val in enumerate(r):
            if re.match(r"^[1-9]\d{0,2}$", col_val):
                num = int(col_val)
                rem_c = [c for c in r[idx + 1:] if c]
                if rem_c:
                    first_col_num = num
                    sl_col_idx = idx
                    break
                
        if first_col_num is not None and sl_col_idx >= 0:
            rem_cells = [c for c in r[sl_col_idx + 1:] if c]
            
            # Filter out duty/vat rows that have a number
            combined_rem = " ".join(rem_cells).lower()
            if re.search(r"\b(duty|vat|tax)\s*\d+%", combined_rem):
                continue
                
            if current_item:
                items.append(current_item)
                current_item = None
                
            desc = ""
            client_po_no = str(first_col_num)
            qty = "1.00"
            uom = "pcs"
            rate = ""
            amount = ""
            
            if len(rem_cells) == 1:
                if re.search(r"[\d,]+\.\d{2}", rem_cells[0]):
                    amount = rem_cells[0]
                else:
                    desc = rem_cells[0]
            elif len(rem_cells) >= 2:
                amount = rem_cells[-1] if re.search(r"[\d,]+\.\d{2}", rem_cells[-1]) else ""
                rate = rem_cells[-2] if len(rem_cells) >= 2 and re.search(r"[\d,]+\.\d{2}", rem_cells[-2]) else ""
                
                remaining = rem_cells[:-2] if (amount and rate) else rem_cells[:-1]
                
                if remaining and re.search(r"\d+(?:\.\d+)?", remaining[-1]):
                    qty_raw = remaining[-1]
                    remaining = remaining[:-1]
                    
                    qty_m = re.search(r"([\d,]+(?:\.\d+)?)", qty_raw)
                    if qty_m:
                        qty = qty_m.group(1).replace(",", "")
                    uom_m = re.search(r"\b(pcs|ea|set|mtr|nos|pkg|m|kg|box|pallet)\b", qty_raw, re.I)
                    if uom_m:
                        uom = uom_m.group(1).lower()
                        
                if remaining and re.match(r"^\d+$", remaining[-1]):
                    client_po_no = remaining[-1]
                    remaining = remaining[:-1]
                    
                desc = " ".join(remaining).strip()
                
            current_item = {
                "sl_no": first_col_num,
                "client_po_item_no": client_po_no,
                "description": desc,
                "qty": qty,
                "uom": uom,
                "rate": rate,
                "amount": amount,
                "remarks": "",
                "make": "",
                "supplier_name": "",
                "supplier_po": "",
                "package_type": "Pallet",
                "package_no": "1"
            }
        else:
            if current_item:
                extra_desc = " ".join(non_empty).strip()
                if extra_desc and not re.search(r"\b(sub total|total|vat|duty)\b", extra_desc.lower()):
                    if current_item["description"]:
                        current_item["description"] = (current_item["description"] + " " + extra_desc).strip()
                    else:
                        current_item["description"] = extra_desc

    if current_item:
        items.append(current_item)

    return items

def _extract_items_from_text(lines):
    """
    Fallback text-based line parser for unstructured or unbordered PDFs.
    """
    items = []
    item_pattern = re.compile(r"^([1-9]\d{0,2})\s+(.+)")
    in_items = False
    current_item = None

    for i, line in enumerate(lines):
        if re.search(r"#\s*Item\s*&\s*Description", line, re.IGNORECASE):
            in_items = True
            continue
        
        if in_items and re.search(r"Sub\s*Total|Standard\s*Rate|Total|VAT", line, re.IGNORECASE):
            if current_item:
                items.append(current_item)
                current_item = None
            break

        if in_items:
            m = item_pattern.match(line)
            if m:
                rest_txt = m.group(2).strip()
                # Check if this line is actually a continuation line that happens to start with a number like "300 LBS" or "316 BALL"
                if re.match(r"^(?:300|150|600|900|1500|2500)\s*(?:LBS|#|PSI|RF|RTJ|SCH)", rest_txt, re.I) or re.match(r"^(?:316|304|316L|304L|A105|LF2|WCB)\s*(?:BALL|STEM|BODY|TRIM|INSERT)", rest_txt, re.I):
                    if current_item:
                        current_item["description"] += " " + line
                        continue

                if current_item:
                    items.append(current_item)
                    current_item = None

                sl_no = m.group(1)
                rest = m.group(2)
                
                num_pattern = re.compile(r"([\d,]+\.\d{2})")
                matches = list(num_pattern.finditer(rest))

                qty = "1.00"
                uom = "pcs"
                description = rest
                rate = ""
                amount = ""

                client_po_item_no = str(sl_no)
                if len(matches) >= 2:
                    rate = matches[-2].group(1)
                    amount = matches[-1].group(1)
                    rate_idx = matches[-2].start()
                    desc_part = rest[:rate_idx].strip()
                    
                    sub_tokens = desc_part.split()
                    for idx in range(len(sub_tokens) - 1, -1, -1):
                        token = sub_tokens[idx]
                        if re.match(r"^\d+(?:\.\d+)?$", token):
                            qty = token
                            if idx + 1 < len(sub_tokens) and not re.match(r"^\d+$", sub_tokens[idx + 1]):
                                uom = sub_tokens[idx + 1]
                            
                            desc_sub = sub_tokens[:idx]
                            if desc_sub and re.match(r"^\d{1,5}$", desc_sub[-1]):
                                client_po_item_no = desc_sub[-1]
                                desc_sub = desc_sub[:-1]
                            
                            description = " ".join(desc_sub)
                            break

                current_item = {
                    "sl_no": int(sl_no),
                    "client_po_item_no": client_po_item_no,
                    "description": description.strip(),
                    "qty": qty if qty else "1.00",
                    "uom": uom if uom else "pcs",
                    "rate": rate,
                    "amount": amount,
                    "heat_number": "",
                    "certificate_number": "",
                    "make": "",
                    "remarks": "",
                    "supplier_name": "",
                    "supplier_po": "",
                    "package_type": "Pallet",
                    "package_no": "1"
                }
            else:
                if current_item:
                    if not re.search(r"Sub\s*Total|Standard\s*Rate|Total|VAT|Page|\d+\s*of\s*\d+", line, re.IGNORECASE):
                        if re.match(r"^[\d\s.,pcsEAea\/]+$", line) and re.search(r"\d+\.\d{2}", line):
                            uom_m = re.search(r"\b(pcs|ea|set|mtr|nos|pkg)\b", line, re.IGNORECASE)
                            if uom_m and not current_item["uom"]:
                                current_item["uom"] = uom_m.group(1).lower()
                            continue

                        clean_line = line
                        uom_match = re.search(r"\b(pcs|ea|set|mtr|nos|pkg)\b", clean_line, re.IGNORECASE)
                        if uom_match:
                            if not current_item["uom"]:
                                current_item["uom"] = uom_match.group(1).lower()
                            clean_line = re.sub(r"\b(pcs|ea|set|mtr|nos|pkg)\b", "", clean_line, flags=re.IGNORECASE).strip()
                        
                        if clean_line and not re.match(r"^\d{4,5}$", clean_line):
                            rem_m = REMARK_PATTERN.search(clean_line)
                            if rem_m:
                                remark_txt = rem_m.group(0).strip()
                                if current_item["remarks"]:
                                    current_item["remarks"] += "; " + remark_txt
                                else:
                                    current_item["remarks"] = remark_txt
                                
                                desc_part = REMARK_PATTERN.sub('', clean_line).rstrip(' -:,;')
                                if desc_part:
                                    current_item["description"] += " " + desc_part
                            else:
                                current_item["description"] += " " + clean_line

    if current_item:
        items.append(current_item)

    return items

def extract_pdf_data(pdf_path_or_file):
    """
    Extracts header metadata (Client, PO Number, MSG Ref) and line items
    from Sales Order / Purchase Order PDFs.
    Auto-detects supplier names from master list and machining instructions.
    """
    full_text = ""
    table_items = []
    try:
        with pdfplumber.open(pdf_path_or_file) as pdf:
            # Extract full text for headers
            for page in pdf.pages:
                page_text = page.extract_text()
                if page_text:
                    full_text += page_text + "\n"
            
            # Primary: Extract structured tables
            table_items = _extract_items_from_tables(pdf)
    except Exception as e:
        print(f"pdfplumber error: {e}")

    if not full_text.strip() and isinstance(pdf_path_or_file, str):
        try:
            with open(pdf_path_or_file, 'r', encoding='utf-8', errors='ignore') as f:
                full_text = f.read()
        except Exception:
            pass

    lines = [line.strip() for line in full_text.splitlines() if line.strip()]

    # --- Header Parsing ---
    client = ""
    po_number = ""
    msg_ref = ""

    so_match = re.search(r"(?:Sales\s*Order\s*#|MSG\s*Ref\s*:?)\s*([A-Z0-9\-_]+)", full_text, re.IGNORECASE)
    if so_match:
        msg_ref = so_match.group(1).strip()

    po_match = re.search(r"(?:Ref\s*#\s*:?|PO\s*(?:Number|#)?\s*:?)\s*([A-Z0-9\-_]+)", full_text, re.IGNORECASE)
    if po_match:
        po_number = po_match.group(1).strip()

    client_match = re.search(r"(?:Bill\s*To|Client\s*:?)\s*\n?([^\n]+)", full_text, re.IGNORECASE)
    if client_match:
        raw_client = client_match.group(1).strip()
        raw_client = re.sub(r"-AED$", "", raw_client, flags=re.IGNORECASE).strip()
        client = raw_client

    # Choose table items if available, else fallback to text parser
    if table_items and len(table_items) > 0:
        items = table_items
    else:
        items = _extract_items_from_text(lines)

    # Post-process items to clean descriptions, auto-detect MAKE & suppliers
    for item in items:
        # Strip header leakage if any (e.g. "1 Client PO Item No")
        item["description"] = re.sub(r"\b\d+\s+Client\s+PO\s+Item\s+No\b", "", item["description"], flags=re.IGNORECASE).strip()
        item["description"] = re.sub(r"\bClient\s+PO\s+Item\s+No\b", "", item["description"], flags=re.IGNORECASE).strip()

        # Strip leading item sequence numbers e.g. "1 ECC. RED..." -> "ECC. RED..."
        item["description"] = strip_leading_item_number(item["description"], item["sl_no"])

        raw_full = item["description"] + " " + item.get("remarks", "")

        # Auto-detect MAKE from code / raw item description if empty
        if not item.get("make"):
            detected_make = auto_detect_make(raw_full)
            if detected_make:
                item["make"] = detected_make
        
        # Check if code or description contains STOCK
        if re.search(r"[-_\s/]?STOCK\b", raw_full, re.IGNORECASE):
            item["supplier_name"] = "STOCK"

        # Check if description itself contains remark patterns
        rem_match = REMARK_PATTERN.search(item["description"])
        if rem_match:
            remark_text = rem_match.group(0).strip()
            cleaned_desc = REMARK_PATTERN.sub('', item["description"]).rstrip(' -:,;')
            item["description"] = cleaned_desc
            if item.get("remarks"):
                if remark_text.lower() not in item["remarks"].lower():
                    item["remarks"] += "; " + remark_text
            else:
                item["remarks"] = remark_text

        # Strip internal product codes e.g. LBL25300RF, LTO156KA350LF2, LUN256KNPTA105N
        item["description"] = strip_internal_product_code(item["description"])

        # Strip leading item number again if revealed after code stripping
        item["description"] = strip_leading_item_number(item["description"], item["sl_no"])

        # Master Supplier Auto-match (if supplier_name is not already set)
        if not item.get("supplier_name"):
            desc_rem = (item["description"] + " " + item.get("remarks", "")).lower()
            EXCLUDED_BRAND_SUPPLIERS = {"mega s.p.a", "metalfar prodotti industriali s.p.a.", "raccortubi s.p.a."}
            for supplier in MASTER_SUPPLIERS:
                if supplier.lower() in EXCLUDED_BRAND_SUPPLIERS:
                    continue
                core_name = re.sub(r"\b(LLC|FZE|FZC|L\.L\.C|PTE|LTD|CO|INC|S\.P\.A|BV)\b", "", supplier, flags=re.IGNORECASE).strip()
                if len(core_name) > 3 and re.search(r"\b" + re.escape(core_name.lower()) + r"\b", desc_rem):
                    item["supplier_name"] = supplier
                    break

        # Ensure default packaging fields
        if not item.get("package_type"):
            item["package_type"] = "Pallet"
        if not item.get("package_no"):
            item["package_no"] = "1"
        if "heat_number" not in item:
            item["heat_number"] = ""
        if "certificate_number" not in item:
            item["certificate_number"] = ""

    # Group items by client_po_item_no to auto-detect split batches (e.g. 14 pcs stock & 26 pcs local market -> 14 OF 40 & 26 OF 40)
    from collections import defaultdict
    client_groups = defaultdict(list)
    for it in items:
        c_no = str(it.get("client_po_item_no", "")).strip()
        if c_no:
            client_groups[c_no].append(it)

    for c_no, grp in client_groups.items():
        if len(grp) > 1:
            total_qty = 0.0
            for it in grp:
                raw_q = str(it.get("qty", "1")).replace(",", "").strip()
                m = re.match(r"^([\d.]+)", raw_q)
                if m:
                    total_qty += float(m.group(1))
            
            fmt_total = int(total_qty) if total_qty.is_integer() else round(total_qty, 2)
            
            for it in grp:
                raw_q = str(it.get("qty", "1")).replace(",", "").strip()
                m = re.match(r"^([\d.]+)", raw_q)
                curr_q = float(m.group(1)) if m else 1.0
                fmt_curr = int(curr_q) if curr_q.is_integer() else round(curr_q, 2)
                it["qty"] = f"{fmt_curr} OF {fmt_total}"
                it["total_qty"] = fmt_total
                it["group_id"] = f"grp-{c_no}"

    return {
        "client": client,
        "po_number": po_number,
        "msg_ref": msg_ref,
        "items": items
    }
