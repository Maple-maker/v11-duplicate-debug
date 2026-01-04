"""DD1750 core - Complete working version."""

import io
import math
import re
from dataclasses import dataclass
from typing import List, Dict

import pdfplumber
from pypdf import PdfReader, PdfWriter
from reportlab.pdfgen import canvas
from reportlab.pdfbase import pdfmetrics
from reportlab.lib.pagesizes import letter


PAGE_W, PAGE_H = letter

# Admin positions - adjust Y values as needed
ADMIN_COORDS = {
    'unit': {'x': 50, 'y': 735},
    'requisition': {'x': 280, 'y': 735},
    'page': {'x': 500, 'y': 735},
    'date': {'x': 50, 'y': 710},
    'order': {'x': 280, 'y': 710},
    'boxes': {'x': 480, 'y': 710},
    'packed_by': {'x': 44, 'y': 115},
    'received_by': {'x': 300, 'y': 115},
}

X_BOX_L, X_BOX_R = 44.0, 88.0
X_CONTENT_L, X_CONTENT_R = 88.0, 365.0
X_UOI_L, X_UOI_R = 365.0, 408.5
X_INIT_L, X_INIT_R = 408.5, 453.5
X_SPARES_L, X_SPARES_R = 453.5, 514.5
X_TOTAL_L, X_TOTAL_R = 514.5, 566.0

Y_TABLE_TOP_LINE = 616.0
Y_TABLE_BOTTOM_LINE = 89.5
ROWS_PER_PAGE = 18
ROW_H = (Y_TABLE_TOP_LINE - Y_TABLE_BOTTOM_LINE) / ROWS_PER_PAGE
PAD_X = 3.0


@dataclass
class BomItem:
    line_no: int
    description: str
    nsn: str
    qty: int


def extract_items_from_pdf(pdf_path: str, start_page: int = 0) -> List[BomItem]:
    items = []
    
    try:
        with pdfplumber.open(pdf_path) as pdf:
            for page in pdf.pages[start_page:]:
                tables = page.extract_tables()
                
                for table in tables:
                    if len(table) < 2:
                        continue
                    
                    header = table[0]
                    lv_idx = desc_idx = mat_idx = qty_idx = None
                    
                    for i, cell in enumerate(header):
                        if cell:
                            text = str(cell).upper()
                            if 'LV' in text or 'LEVEL' in text:
                                lv_idx = i
                            elif 'DESC' in text or 'NOMENCLATURE' in text:
                                desc_idx = i
                            elif 'MATERIAL' in text:
                                mat_idx = i
                            elif 'AUTH' in text or 'QTY' in text or 'QUANTITY' in text:
                                qty_idx = i
                    
                    if lv_idx is None or desc_idx is None:
                        continue
                    
                    for row in table[1:]:
                        if not any(cell for cell in row if cell):
                            continue
                        
                        lv_cell = row[lv_idx] if lv_idx < len(row) else None
                        if not lv_cell or str(lv_cell).strip().upper() != 'B':
                            continue
                        
                        desc_cell = row[desc_idx] if desc_idx < len(row) else None
                        description = ""
                        if desc_cell:
                            lines = str(desc_cell).strip().split('\n')
                            description = lines[1].strip() if len(lines) >= 2 else lines[0].strip()
                            if '(' in description:
                                description = description.split('(')[0].strip()
                            description = re.sub(r'\s+(WTY|ARC|CIIC|UI|SCMC|EA|AY|9K|9G)$', '', description, flags=re.IGNORECASE)
                            description = re.sub(r'\s+', ' ', description).strip()
                        
                        if not description:
                            continue
                        
                        nsn = ""
                        if mat_idx is not None and mat_idx < len(row):
                            mat_cell = row[mat_idx]
                            if mat_cell:
                                match = re.search(r'\b(\d{9})\b', str(mat_cell))
                                if match:
                                    nsn = match.group(1)
                        
                        qty = 1
                        if qty_idx is not None and qty_idx < len(row):
                            qty_cell = row[qty_idx]
                            if qty_cell:
                                qty_text = str(qty_cell).strip()
                                qty_match = re.search(r'(\d+)', qty_text)
                                if qty_match:
                                    qty = int(qty_match.group(1))
                        
                        if qty == 1:
                            parts = description.split()
                            if parts and parts[-1].isdigit():
                                qty = int(parts[-1])
                                description = ' '.join(parts[:-1])
                        
                        items.append(BomItem(len(items) + 1, description[:100], nsn, qty))
    
    except Exception as e:
        print(f"ERROR: {e}")
    
    return items


def generate_dd1750_from_pdf(bom_path, template_path, output_path, start_page=0, admin_data=None):
    if admin_data is None:
        admin_data = {}
    
    items = extract_items_from_pdf(bom_path, start_page)
    
    if not items:
        try:
            reader = PdfReader(template_path)
            writer = PdfWriter()
            writer.add_page(reader.pages[0])
            with open(output_path, 'wb') as f:
                writer.write(f)
        except:
            pass
        return output_path, 0
    
    total_pages = math.ceil(len(items) / ROWS_PER_PAGE)
    writer = PdfWriter()
    
    for page_num in range(total_pages):
        start_idx = page_num * ROWS_PER_PAGE
        end_idx = min((page_num + 1) * ROWS_PER_PAGE, len(items))
        page_items = items[start_idx:end_idx]
        
        packet = io.BytesIO()
        c = canvas.Canvas(packet, pagesize=letter)
        
        first_row_top = Y_TABLE_TOP_LINE - 5.0
        
        for i, item in enumerate(page_items):
            y = first_row_top - (i * ROW_H)
            y_desc = y - 7.0
            y_nsn = y - 12.2
            
            c.setFont("Helvetica", 8)
            c.drawCentredString((X_BOX_L + X_BOX_R) / 2, y_desc, str(item.line_no))
            
            c.setFont("Helvetica", 7)
            desc = item.description[:50] if len(item.description) > 50 else item.description
            c.drawString(X_CONTENT_L + PAD_X, y_desc, desc)
            
            if item.nsn:
                c.setFont("Helvetica", 6)
                c.drawString(X_CONTENT_L + PAD_X, y_nsn, f"NSN: {item.nsn}")
            
            c.setFont("Helvetica", 8)
            c.drawCentredString((X_UOI_L + X_UOI_R) / 2, y_desc, "EA")
            c.drawCentredString((X_INIT_L + X_INIT_R) / 2, y_desc, str(item.qty))
            c.drawCentredString((X_SPARES_L + X_SPARES_R) / 2, y_desc, "0")
            c.drawCentredString((X_TOTAL_L + X_TOTAL_R) / 2, y_desc, str(item.qty))
        
        # Admin fields
        for key, coords in ADMIN_COORDS.items():
            if key == 'page' and total_pages <= 1:
                continue
            
            value = None
            if key == 'requisition' and admin_data.get('requisition_no'):
                value = f"REQ: {admin_data['requisition_no']}"
            elif key == 'order' and admin_data.get('order_no'):
                value = f"ORDER: {admin_data['order_no']}"
            elif key == 'boxes' and admin_data.get('num_boxes'):
                value = f"BOXES: {admin_data['num_boxes']}"
            elif key == 'received_by':
                value = "RECEIVED BY:"
            elif admin_data.get(key):
                value = admin_data[key]
            
            if value:
                c.setFont("Helvetica", 8)
                c.drawString(coords['x'], coords['y'], str(value)[:30])
                
                if key == 'packed_by':
                    c.setFont("Helvetica", 6)
                    c.drawString(coords['x'], coords['y'] - 10, "(Signature)")
                elif key == 'received_by':
                    c.setFont("Helvetica", 6)
                    c.drawString(coords['x'], coords['y'] - 10, "(Signature)")
        
        c.save()
        packet.seek(0)
        
        overlay = PdfReader(packet)
        page = PdfReader(template_path).pages[0]
        page.merge_page(overlay.pages[0])
        writer.add_page(page)
    
    with open(output_path, 'wb') as f:
        writer.write(f)
    
    return output_path, len(items)
