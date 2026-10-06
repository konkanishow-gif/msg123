import os
import io
import re
from datetime import datetime
import openpyxl
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
from openpyxl.drawing.image import Image

def generate_cipl_excel(data, logo_path=None, output_path=None):
    """
    Generates a Commercial Invoice and Packing List (CIPL) Excel workbook with 2 sheets:
    1. COMMERCIAL INVOICE
    2. PACKING LIST
    """
    wb = openpyxl.Workbook()
    
    # Styles
    thin_border = Border(
        left=Side(style='thin', color='000000'),
        right=Side(style='thin', color='000000'),
        top=Side(style='thin', color='000000'),
        bottom=Side(style='thin', color='000000')
    )
    
    company_title_font = Font(name='Calibri', size=15, bold=True, color='C0392B')
    doc_title_font = Font(name='Calibri', size=13, bold=True, color='1B263B')
    meta_title_font = Font(name='Calibri', size=10, bold=True, color='1B263B')
    meta_val_font = Font(name='Calibri', size=10)
    table_header_font = Font(name='Calibri', size=10, bold=True, color='FFFFFF')
    table_header_fill = PatternFill(start_color='1B263B', end_color='1B263B', fill_type='solid')
    data_font = Font(name='Calibri', size=10)
    data_bold_font = Font(name='Calibri', size=10, bold=True)
    zebra_fill = PatternFill(start_color='F8F9FA', end_color='F8F9FA', fill_type='solid')
    total_fill = PatternFill(start_color='E8F4F8', end_color='E8F4F8', fill_type='solid')
    
    client_name = data.get('client', '').strip()
    po_number = data.get('po_number', '').strip()
    msg_ref = data.get('msg_ref', '').strip()
    today_str = datetime.now().strftime('%d-%b-%Y')
    items = data.get('items', [])
    
    # ==========================================
    # SHEET 1: COMMERCIAL INVOICE
    # ==========================================
    ws_inv = wb.active
    ws_inv.title = 'COMMERCIAL INVOICE'
    ws_inv.views.sheetView[0].showGridLines = True
    
    # Header merges
    ws_inv.merge_cells('A1:F1')
    ws_inv['A1'] = 'MSG OILFIELD EQUIPMENT TRADING LLC'
    ws_inv['A1'].font = company_title_font
    ws_inv['A1'].alignment = Alignment(horizontal='left', vertical='center')
    
    ws_inv.merge_cells('A2:F2')
    ws_inv['A2'] = 'Dubai Industrial City (DIC) Phase-1, Dubai, United Arab Emirates | TRN: 100518964000003'
    ws_inv['A2'].font = Font(name='Calibri', size=9, italic=True)
    ws_inv['A2'].alignment = Alignment(horizontal='left', vertical='center')
    
    ws_inv.merge_cells('A3:H3')
    ws_inv['A3'] = 'COMMERCIAL INVOICE'
    ws_inv['A3'].font = doc_title_font
    ws_inv['A3'].alignment = Alignment(horizontal='center', vertical='center')
    ws_inv['A3'].fill = PatternFill(start_color='EAECEE', end_color='EAECEE', fill_type='solid')
    ws_inv.row_dimensions[3].height = 25
    
    # Metadata Box
    ws_inv.merge_cells('A5:D5')
    ws_inv['A5'] = f'BILL TO: {client_name}'
    ws_inv['A5'].font = meta_title_font
    
    ws_inv.merge_cells('A6:D6')
    ws_inv['A6'] = f'CLIENT PO NO: {po_number}'
    ws_inv['A6'].font = meta_val_font
    
    ws_inv.merge_cells('A7:D7')
    ws_inv['A7'] = 'PAYMENT TERMS: As Agreed'
    ws_inv['A7'].font = meta_val_font
    
    ws_inv.merge_cells('E5:H5')
    ws_inv['E5'] = f'INVOICE / SO REF: {msg_ref}'
    ws_inv['E5'].font = meta_title_font
    
    ws_inv.merge_cells('E6:H6')
    ws_inv['E6'] = f'INVOICE DATE: {today_str}'
    ws_inv['E6'].font = meta_val_font
    
    ws_inv.merge_cells('E7:H7')
    ws_inv['E7'] = 'CURRENCY: AED / USD'
    ws_inv['E7'].font = meta_val_font
    
    for r in range(5, 8):
        for c in range(1, 9):
            ws_inv.cell(row=r, column=c).border = thin_border
            
    # Table Headers
    inv_headers = ['SL NO', 'CLIENT PO ITEM NO', 'DESCRIPTION OF GOODS', 'MAKE', 'QTY', 'UOM', 'UNIT RATE', 'TOTAL AMOUNT']
    inv_col_widths = [8, 20, 65, 20, 12, 10, 16, 18]
    
    header_row = 9
    ws_inv.row_dimensions[header_row].height = 25
    for col_idx, (h_text, width) in enumerate(zip(inv_headers, inv_col_widths), start=1):
        cell = ws_inv.cell(row=header_row, column=col_idx)
        cell.value = h_text
        cell.font = table_header_font
        cell.fill = table_header_fill
        cell.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
        cell.border = thin_border
        col_letter = openpyxl.utils.get_column_letter(col_idx)
        ws_inv.column_dimensions[col_letter].width = width
        
    # Table Data
    current_row = 10
    total_invoice_amt = 0.0
    
    for i, it in enumerate(items, start=1):
        sl = it.get('sl_no', i)
        client_po_it = it.get('client_po_item_no', str(sl))
        desc = it.get('description', '')
        make = it.get('make', '')
        qty_str = str(it.get('qty', '1')).replace(',', '').strip()
        
        # Check if qty is in "X OF Y" format
        is_of_format = bool(re.search(r'\b(?:OF|of)\b', qty_str))
        qty_val = 1.0
        if is_of_format:
            m_of = re.search(r'([\d.]+)\s*(?:OF|of)', qty_str, re.I)
            if m_of:
                try:
                    qty_val = float(m_of.group(1))
                except ValueError:
                    qty_val = 1.0
        else:
            try:
                qty_val = float(qty_str) if qty_str else 1.0
            except ValueError:
                qty_val = 1.0
            
        uom = it.get('uom', 'pcs').upper()
        rate_str = str(it.get('rate', '')).replace(',', '').strip()
        try:
            rate_val = float(rate_str) if rate_str else 0.0
        except ValueError:
            rate_val = 0.0
            
        amt_str = str(it.get('amount', '')).replace(',', '').strip()
        try:
            amt_val = float(amt_str) if amt_str else (qty_val * rate_val)
        except ValueError:
            amt_val = qty_val * rate_val
            
        total_invoice_amt += amt_val
        
        row_fill = zebra_fill if (i % 2 == 0) else PatternFill(fill_type=None)
        
        ws_inv.cell(row=current_row, column=1, value=sl).alignment = Alignment(horizontal='center', vertical='center')
        ws_inv.cell(row=current_row, column=2, value=client_po_it).alignment = Alignment(horizontal='center', vertical='center')
        ws_inv.cell(row=current_row, column=3, value=desc).alignment = Alignment(horizontal='left', vertical='center', wrap_text=True)
        ws_inv.cell(row=current_row, column=4, value=make).alignment = Alignment(horizontal='center', vertical='center')
        
        if is_of_format:
            c_qty = ws_inv.cell(row=current_row, column=5, value=qty_str.upper())
            c_qty.alignment = Alignment(horizontal='center', vertical='center')
        else:
            c_qty = ws_inv.cell(row=current_row, column=5, value=qty_val)
            c_qty.alignment = Alignment(horizontal='right', vertical='center')
            c_qty.number_format = '#,##0.00'
        
        ws_inv.cell(row=current_row, column=6, value=uom).alignment = Alignment(horizontal='center', vertical='center')
        
        c_rate = ws_inv.cell(row=current_row, column=7, value=rate_val)
        c_rate.alignment = Alignment(horizontal='right', vertical='center')
        c_rate.number_format = '#,##0.00'
        
        c_amt = ws_inv.cell(row=current_row, column=8, value=amt_val)
        c_amt.alignment = Alignment(horizontal='right', vertical='center')
        c_amt.number_format = '#,##0.00'
        
        for c in range(1, 9):
            ws_inv.cell(row=current_row, column=c).border = thin_border
            ws_inv.cell(row=current_row, column=c).font = data_font
            if i % 2 == 0:
                ws_inv.cell(row=current_row, column=c).fill = row_fill
                
        ws_inv.row_dimensions[current_row].height = 24
        current_row += 1
        
    # Total Row
    ws_inv.merge_cells(f'A{current_row}:G{current_row}')
    ws_inv[f'A{current_row}'] = 'TOTAL COMMERCIAL VALUE'
    ws_inv[f'A{current_row}'].font = data_bold_font
    ws_inv[f'A{current_row}'].alignment = Alignment(horizontal='right', vertical='center')
    
    c_tot = ws_inv.cell(row=current_row, column=8, value=total_invoice_amt)
    c_tot.font = data_bold_font
    c_tot.alignment = Alignment(horizontal='right', vertical='center')
    c_tot.number_format = '#,##0.00'
    
    for c in range(1, 9):
        ws_inv.cell(row=current_row, column=c).border = thin_border
        ws_inv.cell(row=current_row, column=c).fill = total_fill
    ws_inv.row_dimensions[current_row].height = 25
    
    current_row += 2
    ws_inv.cell(row=current_row, column=1, value='Authorized Signature & Company Stamp').font = meta_title_font
    current_row += 4
    ws_inv.merge_cells(f'A{current_row}:D{current_row}')
    ws_inv[f'A{current_row}'] = '_______________________________________'
    
    # ==========================================
    # SHEET 2: PACKING LIST
    # ==========================================
    ws_pk = wb.create_sheet(title='PACKING LIST')
    ws_pk.views.sheetView[0].showGridLines = True
    
    ws_pk.merge_cells('A1:F1')
    ws_pk['A1'] = 'MSG OILFIELD EQUIPMENT TRADING LLC'
    ws_pk['A1'].font = company_title_font
    ws_pk['A1'].alignment = Alignment(horizontal='left', vertical='center')
    
    ws_pk.merge_cells('A2:F2')
    ws_pk['A2'] = 'Dubai Industrial City (DIC) Phase-1, Dubai, UAE'
    ws_pk['A2'].font = Font(name='Calibri', size=9, italic=True)
    
    ws_pk.merge_cells('A3:I3')
    ws_pk['A3'] = 'COMMERCIAL PACKING LIST'
    ws_pk['A3'].font = doc_title_font
    ws_pk['A3'].alignment = Alignment(horizontal='center', vertical='center')
    ws_pk['A3'].fill = PatternFill(start_color='EAECEE', end_color='EAECEE', fill_type='solid')
    ws_pk.row_dimensions[3].height = 25
    
    # Metadata Box
    ws_pk.merge_cells('A5:E5')
    ws_pk['A5'] = f'CLIENT: {client_name}'
    ws_pk['A5'].font = meta_title_font
    
    ws_pk.merge_cells('A6:E6')
    ws_pk['A6'] = f'CLIENT PO NO: {po_number}'
    ws_pk['A6'].font = meta_val_font
    
    ws_pk.merge_cells('F5:I5')
    ws_pk['F5'] = f'SO REF: {msg_ref}'
    ws_pk['F5'].font = meta_title_font
    
    ws_pk.merge_cells('F6:I6')
    ws_pk['F6'] = f'DATE: {today_str}'
    ws_pk['F6'].font = meta_val_font
    
    for r in range(5, 7):
        for c in range(1, 10):
            ws_pk.cell(row=r, column=c).border = thin_border
            
    pk_headers = ['PACKAGE NO / TYPE', 'SL NO', 'CLIENT PO ITEM NO', 'DESCRIPTION OF GOODS', 'QTY', 'UOM', 'NET WT (KG)', 'GROSS WT (KG)', 'DIMENSIONS (CM)']
    pk_col_widths = [20, 8, 20, 60, 10, 10, 14, 14, 18]
    
    header_row_pk = 8
    ws_pk.row_dimensions[header_row_pk].height = 25
    for col_idx, (h_text, width) in enumerate(zip(pk_headers, pk_col_widths), start=1):
        cell = ws_pk.cell(row=header_row_pk, column=col_idx)
        cell.value = h_text
        cell.font = table_header_font
        cell.fill = table_header_fill
        cell.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
        cell.border = thin_border
        col_letter = openpyxl.utils.get_column_letter(col_idx)
        ws_pk.column_dimensions[col_letter].width = width
        
    # Table Data for Packing List
    current_pk_row = 9
    total_qty = 0.0
    packages_seen = set()
    
    for i, it in enumerate(items, start=1):
        pkg_type = it.get('package_type', 'Pallet')
        pkg_no = it.get('package_no', '1')
        full_pkg_label = f'{pkg_type} {pkg_no}'.strip()
        packages_seen.add(full_pkg_label)
        
        sl = it.get('sl_no', i)
        client_po_it = it.get('client_po_item_no', str(sl))
        desc = it.get('description', '')
        qty_str = str(it.get('qty', '1')).replace(',', '').strip()
        
        is_of_format = bool(re.search(r'\b(?:OF|of)\b', qty_str))
        qty_val = 1.0
        if is_of_format:
            m_of = re.search(r'([\d.]+)\s*(?:OF|of)', qty_str, re.I)
            if m_of:
                try:
                    qty_val = float(m_of.group(1))
                except ValueError:
                    qty_val = 1.0
        else:
            try:
                qty_val = float(qty_str) if qty_str else 1.0
            except ValueError:
                qty_val = 1.0
                
        total_qty += qty_val
        uom = it.get('uom', 'pcs').upper()
        
        net_wt = it.get('net_weight', '')
        gross_wt = it.get('gross_weight', '')
        dims = it.get('dimensions', '')
        
        ws_pk.cell(row=current_pk_row, column=1, value=full_pkg_label).alignment = Alignment(horizontal='center', vertical='center')
        ws_pk.cell(row=current_pk_row, column=2, value=sl).alignment = Alignment(horizontal='center', vertical='center')
        ws_pk.cell(row=current_pk_row, column=3, value=client_po_it).alignment = Alignment(horizontal='center', vertical='center')
        ws_pk.cell(row=current_pk_row, column=4, value=desc).alignment = Alignment(horizontal='left', vertical='center', wrap_text=True)
        
        if is_of_format:
            c_q = ws_pk.cell(row=current_pk_row, column=5, value=qty_str.upper())
            c_q.alignment = Alignment(horizontal='center', vertical='center')
        else:
            c_q = ws_pk.cell(row=current_pk_row, column=5, value=qty_val)
            c_q.alignment = Alignment(horizontal='right', vertical='center')
            c_q.number_format = '#,##0.00'
        
        ws_pk.cell(row=current_pk_row, column=6, value=uom).alignment = Alignment(horizontal='center', vertical='center')
        ws_pk.cell(row=current_pk_row, column=7, value=net_wt).alignment = Alignment(horizontal='center', vertical='center')
        ws_pk.cell(row=current_pk_row, column=8, value=gross_wt).alignment = Alignment(horizontal='center', vertical='center')
        ws_pk.cell(row=current_pk_row, column=9, value=dims).alignment = Alignment(horizontal='center', vertical='center')
        
        for c in range(1, 10):
            ws_pk.cell(row=current_pk_row, column=c).border = thin_border
            ws_pk.cell(row=current_pk_row, column=c).font = data_font
            if i % 2 == 0:
                ws_pk.cell(row=current_pk_row, column=c).fill = zebra_fill
                
        ws_pk.row_dimensions[current_pk_row].height = 24
        current_pk_row += 1
        
    # Total Packing Summary Row
    ws_pk.merge_cells(f'A{current_pk_row}:D{current_pk_row}')
    ws_pk[f'A{current_pk_row}'] = f'TOTAL: {len(packages_seen)} PACKAGE(S) | {len(items)} LINE ITEMS'
    ws_pk[f'A{current_pk_row}'].font = data_bold_font
    ws_pk[f'A{current_pk_row}'].alignment = Alignment(horizontal='left', vertical='center')
    
    c_tot_q = ws_pk.cell(row=current_pk_row, column=5, value=total_qty)
    c_tot_q.font = data_bold_font
    c_tot_q.alignment = Alignment(horizontal='right', vertical='center')
    c_tot_q.number_format = '#,##0.00'
    
    for c in range(1, 10):
        ws_pk.cell(row=current_pk_row, column=c).border = thin_border
        ws_pk.cell(row=current_pk_row, column=c).fill = total_fill
    ws_pk.row_dimensions[current_pk_row].height = 25
    
    current_pk_row += 2
    ws_pk.cell(row=current_pk_row, column=1, value='Authorized Signature & Warehouse Stamp').font = meta_title_font
    current_pk_row += 4
    ws_pk.merge_cells(f'A{current_pk_row}:D{current_pk_row}')
    ws_pk[f'A{current_pk_row}'] = '_______________________________________'
    
    if output_path:
        wb.save(output_path)
        return output_path
    else:
        out = io.BytesIO()
        wb.save(out)
        out.seek(0)
        return out
