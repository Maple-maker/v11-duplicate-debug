"""DD1750 core - Bulletproof file generation."""

import io
import math
import re
import sys
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
    """Extract items from BOM PDF."""
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
                    
                    for i, cell in enumerate(header):
                        if cell:
                            text = str(cell).upper()
                            if 'LV' in text or 'LEVEL' in text:
                                lv_idx = i
                            elif 'DESC' in text:
                                desc_idx = i
                            elif 'MATERIAL' in text:
                                mat_idx = i
                            elif 'AUTH' in text and 'QTY' in text:
                                auth_idx = i
                    
                    if lv_idx == -1 or desc_idx == -1:
                        print("DEBUG: Skipping table - no LV or DESC")
                        continue
                    
                    for row in table[1:]:
                        if not any(cell for cell in row if cell):
                            continue
                        
                        lv_cell = row[lv_idx]
                        if not lv_cell or str(lv_cell).strip().upper() != 'B':
                            continue
                        
                        desc_cell = row[desc_idx]
                        
                        # Smart description extraction
                        description = ""
                        if desc_cell:
                            lines = str(desc_cell).strip().split('\n')
                            # Pick best line (handles EPP single-line vs BCP multi-line)
                            if len(lines) >= 2:
                                description = lines[1].strip()
                            else:
                                description = lines[0].strip()
                            
                            # Cleanup
                            if '(' in description:
                                description = description.split('(')[0].strip()
                            
                            # Remove codes
                            description = re.sub(r'\s+(WTY|ARC|CIIC|UI|SCMC|EA|AY|9K|9G)$', '', description, flags=re.IGNORECASE)
                            description = re.sub(r'\s+', ' ', description).strip()
                        
                        if not description or len(description) < 3:
                            print(f"DEBUG: Skipping item (invalid description)")
                            continue
                        
                        nsn = ""
                        if mat_idx > -1 and mat_idx < len(row):
                            mat_cell = row[mat_idx]
                            if mat_cell:
                                match = re.search(r'\b(\d{9})\b', str(mat_cell))
                                if match:
                                    nsn = match.group(1)
                        
                        qty = 1
                        if auth_idx > -1 and auth_idx < len(row):
                            qty_cell = row[auth_idx]
                            if qty_cell:
                                match = re.search(r'(\d+)', str(qty_cell))
                                if match:
                                    qty = int(match.group(1))
                        
                        items.append(BomItem(len(items) + 1, description[:100], nsn, qty))
                        print(f"DEBUG: Added item {len(items)}")
    
    except Exception as e:
        print(f"ERROR: {e}", file=sys.stderr)
        import traceback
        traceback.print_exc()
        return []
    
    return items


def generate_dd1750_from_pdf(bom_path: str, template_path: str, out_path: str):
    """Generate DD1750 with robust file writing."""
    items = extract_items_from_pdf(bom_path)
    
    print(f"DEBUG: Extracted {len(items)} items")
    
    if not items:
        print("DEBUG: No items, writing blank template")
        try:
            reader = PdfReader(template_path)
            writer = PdfWriter()
            writer.add_page(reader.pages[0])
            with open(out_path, 'wb') as f:
                writer.write(f)
        except Exception as e:
            print(f"ERROR writing blank: {e}")
        return out_path, 0
    
    # Write to absolute path (not temp) to ensure persistence
    writer = PdfWriter()
    template = PdfReader(template_path)
    
    # Calculate pages
    total_pages = math.ceil(len(items) / ROWS_PER_PAGE)
    
    print(f"DEBUG: Generating {total_pages} pages")
    
    for page_num in range(total_pages):
        start_idx = page_num * ROWS_PER_PAGE
        end_idx = min((page_num + 1) * ROWS_PER_PAGE, len(items))
        page_items = items[start_idx:end_idx]
        
        print(f"DEBUG: Page {page_num} has {len(page_items)} items")
        
        # Create overlay
        packet = io.BytesIO()
        c = canvas.Canvas(packet, pagesize=letter)
        first_row = Y_TABLE_TOP - 5.0
        
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
        
        c.save()
        packet.seek(0)
        
        # Merge overlay
        overlay = PdfReader(packet)
        page = template.pages[page_num]
        page.merge_page(overlay.pages[0])
        writer.add_page(page)
        print(f"DEBUG: Added page {page_num}")
    
    # Write to absolute path (CRITICAL)
    print(f"DEBUG: Writing to {out_path}")
    
    # Explicitly flush before returning
    with open(out_path, 'wb') as f:
        writer.write(f)
    f.flush()
    os.fsync(f.fileno())
    
    print(f"DEBUG: Final file check - Exists: {os.path.exists(out_path)}, Size: {os.path.getsize(out_path)}")
    
    return out_path, len(items)
