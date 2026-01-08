"""DD1750 core - Super Robust Version."""

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


def get_description_text(desc_cell) -> str:
    """Extract description from a table cell using multiple strategies."""
    if not desc_cell:
        return ""
    
    # Convert to string and clean
    text = str(desc_cell).strip()
    
    # Strategy 1: Smart line selection (handles EPP/BCP styles)
    lines = text.split('\n')
    candidates = []
    
    for i, ln in enumerate(lines):
        ln = ln.strip()
        if not ln:
            continue
        # Skip single character codes
        if len(ln) <= 3 and ln.isupper():
            continue
        # Filter out obvious codes
        if re.match(r'^(WTY|ARC|CIIC|UI|SCMC|EA|AY|9K|9G|TMY|SMD|ECC|HEADER|PART)$', ln.upper()):
            continue
        # Prefer longer lines, lines with mixed case
        if len(ln) > 3 or (any(c.islower() for c in ln) and any(c.isupper() for c in ln)):
            candidates.append(ln)
        # Prefer lines containing numbers or parentheses (likely descriptions)
        elif re.search(r'\d', ln) or '(' in ln:
            candidates.append(ln)
    
    # Select best candidate
    if candidates:
        # Sort by length descending, then by specificity
        candidates.sort(key=lambda x: (-len(x), not x.isupper()), reverse=True)
        description = candidates[0].strip()
        
        # Clean up the selected description
        # Remove parentheses and everything inside (some BOMs have NSN in desc)
        description = re.sub(r'\(.*$', '', description).strip()
        # Remove trailing codes
        description = re.sub(r'\s+(WTY|ARC|CIIC|UI|SCMC|EA|AY|9K|9G|TMY|SMD|ECC|OH|QTY|AUTH|MAT|NSN|SERIAL)$', '', description, flags=re.IGNORECASE)
        # Normalize spaces
        description = re.sub(r'\s+', ' ', description).strip()
        
        if len(description) < 3:
            # Fallback: Use the first line of the cell (for EPP style BOMs where header might be on line 1)
            description = lines[0].strip() if lines else text
    elif len(text) > 3:
        # Use the raw text if it looks substantial
        description = text.strip()[:100]
    else:
        # Fallback
        description = text.strip()[:100]
    
    return description


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
                            elif 'OH' in text and 'QTY' in text:
                                auth_idx = i
                            elif 'AUTH' in text and 'QTY' in text:
                                auth_idx = i
                    
                    if lv_idx == -1 or desc_idx == -1:
                        continue
                    
                    for row in table[1:]:
                        if not any(cell for cell in row if cell):
                            continue
                        
                        # Check Level
                        lv_cell = row[lv_idx] if lv_idx < len(row) else None
                        if not lv_cell:
                            continue
                            
                        if str(lv_cell).strip().upper() != 'B' and str(lv_cell).strip().upper() != 'B9':
                            continue
                        
                        # Get description using robust helper
                        desc_cell = row[desc_idx] if desc_idx < len(row) else None
                        description = get_description_text(desc_cell)
                        
                        if not description:
                            continue
                        
                        # Get NSN
                        nsn = ""
                        if mat_idx > -1 and mat_idx < len(row):
                            mat_cell = row[mat_idx]
                            if mat_cell:
                                match = re.search(r'\b(\d{9})\b', str(mat_cell))
                                if match:
                                    nsn = match.group(1)
                        
                        # Get Quantity (OH QTY or AUTH QTY)
                        qty = 1
                        if auth_idx > -1 and auth_idx < len(row):
                            qty_cell = row[auth_idx]
                            if qty_cell:
                                match = re.search(r'(\d+)', str(qty_cell))
                                if match:
                                    qty = int(match.group(1))
                        elif oh_qty_idx > -1 and oh_qty_idx < len(row):
                            qty_cell = row[oh_qty_idx]
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


def generate_dd1750_from_pdf(bom_path: str, template_path: str, out_path: str, start_page: int = 0):
    """Generate DD1750."""
    items = extract_items_from_pdf(bom_path, start_page)
    
    print(f"\nItems found: {len(items)}")
    
    if not items:
        # Write empty template to ensure user gets a file
        try:
            reader = PdfReader(template_path)
            writer = PdfWriter()
            writer.add_page(reader.pages[0])
            with open(out_path, 'wb') as f:
                writer.write(f)
            print(f"Wrote empty template to {out_path}")
        except Exception as e:
            print(f"ERROR writing empty template: {e}")
            pass
        return out_path, 0
    
    total_pages = math.ceil(len(items) / ROWS_PER_PAGE)
    writer = PdfWriter()
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
            y_desc = y - 7.0
            y_nsn = y - 12.2
            
            c.setFont("Helvetica", 8)
            c.drawCentredString((X_BOX_L + X_BOX_R) / 2, y_desc, str(item.line_no))
            
            c.setFont("Helvetica", 7)
            c.drawString(X_CONTENT_L + PAD_X, y_desc, item.description[:50])
            
            if item.nsn:
                c.setFont("Helvetica", 6)
                c.drawString(X_CONTENT_L + PAD_X, y_nsn, f"NSN: {item.nsn}")
            
            c.setFont("Helvetica", 8)
            c.drawCentredString((X_UOI_L + X_UOI_R) / 2, y_desc, "EA")
            c.drawCentredString((X_INIT_L + X_INIT_R) / 2, y_desc, str(item.qty))
            c.drawCentredString((X_SPARES_L + X_SPARES_R) / 2, y_desc, "0")
            c.drawCentredString((X_TOTAL_L + X_TOTAL_R) / 2, y_desc, str(item.qty))
        
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
    
    # Write to file with verification
    try:
        with open(out_path, 'wb') as f:
            writer.write(f)
            print(f"DEBUG: Wrote {out_path}")
            sys.stdout.flush()
    except Exception as e:
        print(f"ERROR writing PDF: {e}", file=sys.stderr)
        import traceback
        traceback.print_exc()
        # Try fallback
        try:
            writer_fallback = PdfWriter()
            writer_fallback.add_page(template.pages[0])
            with open(out_path, 'wb') as f:
                writer_fallback.write(f)
            print(f"FALLBACK: Wrote simple template to {out_path}")
            sys.stdout.flush()
        except:
            pass
    
    return out_path, len(items)
