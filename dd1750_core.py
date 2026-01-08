"""DD1750 core - Bulletproof Version."""

import io
import math
import re
from dataclasses import dataclass
from typing import List

import pdfplumber
from pypdf import PdfReader, PdfWriter
from reportlab.pdfgen import canvas
from reportlab.pdfbase import pdfmetrics
from reportlab.lib.pagesizes import letter


PAGE_W, PAGE_H = letter

# Column positions
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
                    lv_idx = desc_idx = mat_idx = auth_idx = -1
                    oh_qty_idx = -1
                    
                    # Identify columns - EXPAND to find OH QTY
                    for i, cell in enumerate(header):
                        if cell:
                            text = str(cell).upper()
                            if 'LV' in text or 'LEVEL' in text:
                                lv_idx = i
                            elif 'DESC' in text:
                                desc_idx = i
                            elif 'MATERIAL' in text:
                                mat_idx = i
                            elif 'OH' in text and 'QTY' in text:
                                oh_qty_idx = i
                            elif 'AUTH' in text and 'QTY' in text:
                                auth_idx = i
                    
                    if lv_idx == -1 or desc_idx == -1:
                        continue
                    
                    for row in table[1:]:
                        if not any(cell for cell in row if cell):
                            continue
                        
                        # Check Level B (allow B, B9, etc.)
                        lv_cell = row[lv_idx]
                        if not lv_cell or str(lv_cell).strip().upper()[:1] != 'B':
                            continue
                        
                        # Extract Description - Use RAW content
                        desc_cell = row[desc_idx]
                        description = ""
                        if desc_cell:
                            text = str(desc_cell).strip()
                            
                            # Remove trailing codes like (WTY ARC, etc.)
                            text = re.sub(r'\s+(WTY|ARC|CIIC|UI|SCMC|EA|AY|9K|9G)$', '', text, flags=re.IGNORECASE)
                            text = re.sub(r'\s+', ' ', text).strip()
                            
                            # Remove parentheses and content after them (NSN)
                            text = re.sub(r'\(.*$', '', text).strip()
                            
                            if not text:
                                continue
                            
                            description = text[:100]
                        
                        # Extract NSN
                        nsn = ""
                        if mat_idx > -1 and mat_idx < len(row):
                            mat_cell = row[mat_idx]
                            if mat_cell:
                                match = re.search(r'\b(\d{9})\b', str(mat_cell))
                                if match:
                                    nsn = match.group(1)
                        
                        # Extract Quantity - PREFER OH QTY, FALLBACK TO AUTH QTY
                        qty = 1
                        if oh_qty_idx > -1 and oh_qty_idx < len(row):
                            qty_cell = row[oh_qty_idx]
                            if qty_cell:
                                try:
                                    qty = int(str(qty_cell).strip())
                                except:
                                    qty = 1
                        elif auth_idx > -1 and auth_idx < len(row):
                            qty_cell = row[auth_idx]
                            if qty_cell:
                                try:
                                    qty = int(str(qty_cell).strip())
                                except:
                                    qty = 1
                        
                        items.append(BomItem(len(items) + 1, description, nsn, qty))
    
    except Exception as e:
        print(f"ERROR: {e}")
        import traceback
        traceback.print_exc()
        return []
    
    return items


def generate_dd1750_from_pdf(bom_path: str, template_path: str, out_path: str):
    """Generate DD1750 with items overlayed on template."""
    items = extract_items_from_pdf(bom_path)
    
    if not items:
        # If no items, return a minimal PDF
        reader = PdfReader(template_path)
        writer = PdfWriter()
        writer.add_page(reader.pages[0])
        with open(out_path, 'wb') as f:
            writer.write(f)
        return out_path, 0
    
    total_pages = math.ceil(len(items) / ROWS_PER_PAGE)
    writer = PdfWriter()
    
    # Read template
    template = PdfReader(template_path)
    
    for page_num in range(total_pages):
        start_idx = page_num * ROWS_PER_PAGE
        end_idx = min((page_num + 1) * ROWS_PER_PAGE, len(items))
        page_items = items[start_idx:end_idx]
        
        # Create overlay
        packet = io.BytesIO()
        c = canvas.Canvas(packet, pagesize=letter)
        first_row = Y_TABLE_TOP - 5.0
        
        for i, item in enumerate(page_items):
            y = first_row - (i * ROW_H)
            
            c.setFont("Helvetica", 8)
            c.drawCentredString((X_BOX_L + X_BOX_R) / 2, y - 7, str(item.line_no))
            
            c.setFont("Helvetica", 7)
            c.drawString(X_CONTENT_L + PAD_X, y - 7, item.description[:50])
            
            if item.nsn:
                c.setFont("Helvetica", 6)
                c.drawString(X_CONTENT_L + PAD_X, y - 12, f"NSN: {item.nsn}")
            
            c.setFont("Helvetica", 8)
            c.drawCentredString((X_UOI_L + X_UOI_R) / 2, y - 7, "EA")
            c.drawCentredString((X_INIT_L + X_INIT_R) / 2, y - 7, str(item.qty))
            c.drawCentredString((X_SPARES_L + X_SPARES_R) / 2, y - 7, "0")
            c.drawCentredString((X_TOTAL_L + X_TOTAL_R) / 2, y - 7, str(item.qty))
        
        c.save()
        packet.seek(0)
        
        # Merge overlay
        overlay = PdfReader(packet)
        
        # Get template page (use first page for all to ensure consistency)
        if page_num < len(template.pages):
            page = template.pages[page_num]
        else:
            page = template.pages[0]
        
        page.merge_page(overlay.pages[0])
        writer.add_page(page)
    
    # Write to file
    with open(out_path, 'wb') as f:
        writer.write(f)
    
    return out_path, len(items)
