import os
import json
import time
import uuid
from datetime import datetime

HISTORY_FILE = os.path.join(os.path.dirname(__file__), 'static', 'history', 'history.json')

def _ensure_history_dir():
    os.makedirs(os.path.dirname(HISTORY_FILE), exist_ok=True)
    if not os.path.exists(HISTORY_FILE):
        with open(HISTORY_FILE, 'w', encoding='utf-8') as f:
            json.dump([], f)

def get_all_sheets(search_query=""):
    """
    Returns unlimited historical worksheet records, sorted newest first.
    Filters by search_query (Client Name, PO Number, MSG Ref, Date, Item Keywords).
    """
    _ensure_history_dir()
    try:
        with open(HISTORY_FILE, 'r', encoding='utf-8') as f:
            sheets = json.load(f)
    except Exception:
        sheets = []

    # Sort newest first
    sheets.sort(key=lambda x: x.get('timestamp', ''), reverse=True)

    if not search_query or not search_query.strip():
        return sheets

    q = search_query.strip().lower()

    filtered = []
    for s in sheets:
        client = str(s.get('client', '')).lower()
        po = str(s.get('po_number', '')).lower()
        msg_ref = str(s.get('msg_ref', '')).lower()
        timestamp = str(s.get('formatted_date', '')).lower()
        
        # Search inside items text as well
        item_text = " ".join(
            str(it.get('description', '')) + " " + str(it.get('remarks', '')) + " " + str(it.get('make', ''))
            for it in s.get('items', [])
        ).lower()

        if q in client or q in po or q in msg_ref or q in timestamp or q in item_text:
            filtered.append(s)

    return filtered

def get_sheet_by_id(sheet_id):
    """
    Retrieves a single saved worksheet record by unique ID.
    """
    sheets = get_all_sheets()
    for s in sheets:
        if s.get('id') == sheet_id:
            return s
    return None

def save_sheet_snapshot(data, file_urls=None):
    """
    Saves or updates a worksheet state snapshot in the history database.
    data: JSON payload with client, po_number, msg_ref, items, filter_machined_only
    file_urls: dict with client_excel_url, stores_excel_url, supplier_pdf_zip_url
    """
    _ensure_history_dir()
    sheets = get_all_sheets()

    msg_ref = data.get("msg_ref", "SO").strip()
    po_number = data.get("po_number", "").strip()
    client = data.get("client", "").strip()
    items = data.get("items", [])

    # Create unique ID if not present
    existing_id = data.get("id")
    sheet_id = existing_id if existing_id else f"sheet_{datetime.now().strftime('%Y%m%d_%H%M%S')}_{uuid.uuid4().hex[:6]}"

    formatted_date = datetime.now().strftime('%d %b %Y, %I:%M %p')

    record = {
        "id": sheet_id,
        "timestamp": datetime.now().isoformat(),
        "formatted_date": formatted_date,
        "client": client,
        "po_number": po_number,
        "msg_ref": msg_ref,
        "item_count": len(items),
        "filter_machined_only": bool(data.get("filter_machined_only", True)),
        "items": items,
        "file_urls": file_urls or {}
    }

    # Remove existing record if updating
    sheets = [s for s in sheets if s.get('id') != sheet_id]
    sheets.insert(0, record)

    with open(HISTORY_FILE, 'w', encoding='utf-8') as f:
        json.dump(sheets, f, indent=2)

    return record

def delete_sheet_snapshot(sheet_id):
    """
    Deletes a historical sheet record by ID.
    """
    _ensure_history_dir()
    sheets = get_all_sheets()
    filtered = [s for s in sheets if s.get('id') != sheet_id]
    with open(HISTORY_FILE, 'w', encoding='utf-8') as f:
        json.dump(filtered, f, indent=2)
    return True
