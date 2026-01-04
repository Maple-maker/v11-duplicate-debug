"""DD1750 core - Fixed admin positions and date format."""

import io
import math
import re
from dataclasses import dataclass
from typing import List, Dict

import pdfplumber
from pypdf import PdfReader, PdfWriter
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import letter


PAGE_W, PAGE_H = letter

# === ADMIN FIELD POSITIONS (ADJUST AS NEEDED) ===
# Look at your template and adjust these X, Y values
ADMIN_POS = {
    # Top section labels - find these labels and position values AFTER them
    'UNIT': {'label_x': 44, 'value_x': 150, 'y': 745},
    'REQUISITION': {'label_x': 250, 'value_x': 380, 'y': 745},
    'PAGE': {'label_x': 500, 'value_x': 535, 'y': 745},
    'DATE': {'label_x': 44, 'value_x': 150, 'y': 715},
    'ORDER': {'label_x': 250, 'value_x': 380, 'y': 715},
    'BOXES': {'label_x': 450, 'value_x': 550, 'y': 715},
    # Bottom section
    'PACKED': {'label_x': 44, 'value_x': 150, 'y': 120},
    'RECEIVED': {'label_x': 300, 'value_x': 420, 'y': 120},
}

# Table positions
X_BOX_L, X_BOX_R = 44.0, 88.0
X_CONTENT_L, X_CONTENT_R = 88.0, 365.0
X_UOI_L, X_UOI_R = 365.0, 408.5
X_INIT_L, X_INIT_R = 408.5, 453.5
X_SPARES_L, X_SPARES_R = 453.5, 514.5
X_TOTAL_L, X_TOTAL_R = 514.5, 566.0

Y_TABLE_TOP = 616.0
Y_TABLE_BOTTOM = 89.5
ROWS_PER_PAGE = 18
ROW_H = (Y_TABLE_TOP - Y_TABLE_BOTTOM) / ROWS_PER_PAGE


@dataclass
class BomItem:
    line_no: int
    description: str
    nsn: str
    qty: int


def format_military_date(date_str: str) -> str:
    """Convert date to military format DDMONYYYY (ex: 04JAN2026)."""
    if not date_str:
        return ""
    
    try:
        # Try parsing various date formats
        for fmt in ['%Y-%m-%d', '%m/%d/%Y', '%d/%m/%Y', '%Y/%m/%d']:
            try:
                dt = datetime.strptime(date_str, fmt)
                return dt.strftime('%d%b%Y').upper()
            except:
                continue
        
        # If already in correct format, return as-is
        if re.match(r'^\d{2}[A-Z]{3}\d{4}$', date_str.upper()):
            return date_str.upper()
        
        return date_str
    except:
        return date_str


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
                    lv_idx = desc_idx = mat_idx = auth_qty_idx = None
                    
                    for i, cell in enumerate(header):
                        if cell:
                            text = str(cell).upper()
                            if 'LV' in text or 'LEVEL' in text:
                                lv_idx = i
                            elif 'DESC' in text or 'NOMENCLATURE' in text:
                                desc_idx = i
                            elif 'MATERIAL' in text:
                                mat_idx = i
                            elif 'AUTH' in text and 'QTY' in text:
                                auth_qty_idx = i
                    
                    if lv_idx is None or desc_idx is None:
                        continue
                    
                    for row in table[1:]:
                        if not any(cell for cell in row if cell):
                            continue
                        
                        lv_cell = row[lv_idx] if lv_idx < len(row) else None
                        if not lv_cell or str(lv_cell).strip().upper() != 'B':
                            continue
                        
                        # Get description - second line
                        desc_cell = row[desc_idx] if desc_idx < len(row) else None
                        description = ""
                        if desc_cell:
                            lines = str(desc_cell).strip().split('\n')
                            description = lines[1].strip() if len(lines) >= 2 else lines[0].strip()
                            if '(' in description:
                                description = description.split('(')[0].strip()
                        
                        if not description:
                            continue
                        
                        # Get NSN
                        nsn = ""
                        if mat_idx is not None and mat_idx < len(row):
                            mat_cell = row[mat_idx]
                            if mat_cell:
                                match = re.search(r'\b(\d{9})\b', str(mat_cell))
                                if match:
                                    nsn = match.group(1)
                        
                        # Get QTY from AUTH QTY column
                        qty = 1
                        if auth_qty_idx is not None and auth_qty_idx < len(row):
                            qty_cell = row[auth_qty_idx]
                            if qty_cell:
                                qty_text = str(qty_cell).strip()
                                match = re.search(r'(\d+)', qty_text)
                                if match:
                                    qty = int(match.group(1))
                        
                        items.append(BomItem(len(items) + 1, description.strip(), nsn, qty))
    
    except Exception as e:
        print(f"ERROR: {e}")
    
    return items


def generate_dd1750_from_pdf(bom_path: str, template_path: str, output_path: str,
                            start_page: int = 0, admin_data: Dict = None) -> tuple:
    if admin_data is None:
        admin_data = {}
    
    # Format date to military format
    if admin_data.get('date'):
        admin_data['date'] = format_military_date(admin_data['date'])
    
    items = extract_items_from_pdf(bom_path, start_page)
    print(f"Items: {len(items)}, Date format: {admin_data.get('date')}")
    
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
        first_row = Y_TABLE_TOP - 5.0
        
        # Draw items
        for i, item in enumerate(page_items):
            y = first_row - (i * ROW_H)
            
            c.setFont("Helvetica", 8)
            c.drawCentredString(66, y - 7, str(item.line_no))
            
            c.setFont("Helvetica", 7)
            c.drawString(92, y - 7, item.description[:50])
            
            if item.nsn:
                c.setFont("Helvetica", 6)
                c.drawString(92, y - 12, f"NSN: {item.nsn}")
            
            c.setFont("Helvetica", 8)
            c.drawCentredString(386, y - 7, "EA")
            c.drawCentredString(431, y - 7, str(item.qty))
            c.drawCentredString(484, y - 7, "0")
            c.drawCentredString(540, y - 7, str(item.qty))
        
        # === DRAW ADMIN FIELDS ===
        c.setFont("Helvetica", 10)
        
        # UNIT (top left)
        if admin_data.get('unit'):
            c.drawString(ADMIN_POS['UNIT']['value_x'], ADMIN_POS['UNIT']['y'], admin_data['unit'][:25])
        
        # REQUISITION NO.
        if admin_data.get('requisition_no'):
            c.drawString(ADMIN_POS['REQUISITION']['value_x'], ADMIN_POS['REQUISITION']['y'], 
                        f"REQ: {admin_data['requisition_no']}")
        
        # PAGE
        if total_pages > 1:
            c.setFont("Helvetica", 8)
            c.drawString(ADMIN_POS['PAGE']['value_x'], ADMIN_POS['PAGE']['y'], 
                        f"{page_num + 1}/{total_pages}")
        
        # DATE (military format)
        if admin_data.get('date'):
            c.setFont("Helvetica", 10)
            c.drawString(ADMIN_POS['DATE']['value_x'], ADMIN_POS['DATE']['y'], admin_data['date'])
        
        # ORDER NO.
        if admin_data.get('order_no'):
            c.drawString(ADMIN_POS['ORDER']['value_x'], ADMIN_POS['ORDER']['y'],
                        f"ORDER: {admin_data['order_no']}")
        
        # TOTAL NO. OF BOXES
        if admin_data.get('num_boxes'):
            c.drawString(ADMIN_POS['BOXES']['value_x'], ADMIN_POS['BOXES']['y'], admin_data['num_boxes'])
        
        # PACKED BY (bottom section - every page)
        if admin_data.get('packed_by'):
            c.setFont("Helvetica", 10)
            c.drawString(ADMIN_POS['PACKED']['value_x'], ADMIN_POS['PACKED']['y'], admin_data['packed_by'])
            c.setFont("Helvetica", 8)
            c.drawString(ADMIN_POS['PACKED']['value_x'], ADMIN_POS['PACKED']['y'] - 10, "(Signature)")
        
        # RECEIVED BY (bottom section - every page)
        c.setFont("Helvetica", 10)
        c.drawString(ADMIN_POS['RECEIVED']['value_x'], ADMIN_POS['RECEIVED']['y'], "RECEIVED BY:")
        c.setFont("Helvetica", 8)
        c.drawString(ADMIN_POS['RECEIVED']['value_x'], ADMIN_POS['RECEIVED']['y'] - 10, "(Signature)")
        
        c.save()
        packet.seek(0)
        
        overlay = PdfReader(packet)
        page = PdfReader(template_path).pages[0]
        page.merge_page(overlay.pages[0])
        writer.add_page(page)
    
    with open(output_path, 'wb') as f:
        writer.write(f)
    
    return output_path, len(items)
