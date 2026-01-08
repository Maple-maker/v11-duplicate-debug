"""DD1750 core - Robust, OH QTY优先."""

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


PAGE_W, PAGE_H = 612.0, 792.0

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
        print(f"DEBUG: Attempting to extract from {pdf_path}")
        
        with pdfplumber.open(pdf_path) as pdf:
            for page in pdf.pages[start_page:]:
                tables = page.extract_tables()
                
                for table in tables:
                    if len(table) < 2:
                        continue
                    
                    header = table[0]
                    lv_idx = desc_idx = mat_idx = -1
                    oh_qty_idx = auth_qty_idx = -1
                    
                    # Column Identification - Prioritize OH QTY
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
                                auth_qty_idx = i
                    
                    print(f"DEBUG: Columns - LV:{lv_idx}, DESC:{desc_idx}, MAT:{mat_idx}, OH QTY:{oh_qty_idx}, AUTH QTY:{auth_qty_idx}")
                    
                    if lv_idx == -1 or desc_idx == -1:
                        print("DEBUG: Skipping table - no LV or DESC column")
                        continue
                    
                    for row in table[1:]:
                        if not any(cell for cell in row if cell):
                            continue
                        
                        # Check Level B
                        lv_cell = row[lv_idx] if lv_idx < len(row) else None
                        if not lv_cell or str(lv_cell).strip().upper() != 'B':
                            continue
                        
                        # Extract Description - Use entire raw content of the cell
                        desc_cell = row[desc_idx] if desc_idx < len(row) else None
                        description = ""
                        if desc_cell:
                            # Get entire content, replace newlines with spaces to be safe
                            # Then clean minimally
                            raw_text = str(desc_cell).strip()
                            
                            # Replace newlines with spaces first
                            description = raw_text.replace('\n', ' ')
                            
                            # Remove trailing garbage (codes, extra spaces)
                            description = re.sub(r'\s+(WTY|ARC|CIIC|UI|SCMC|EA|AY|9K|9G)$', '', description, flags=re.IGNORECASE)
                            description = re.sub(r'\s+', ' ', description).strip()
                            
                            # Remove parentheses and content after them (some BOMs have NSN in desc)
                            if '(' in description:
                                description = description.split('(')[0].strip()
                            
                            # Fallback: if empty, use empty string
                            if not description:
                                description = ""
                            
                        print(f"DEBUG: Description extracted: '{description[:40]}...'")
                        
                        if not description:
                            print(f"DEBUG: No valid description found for row")
                            continue
                        
                        # Extract NSN
                        nsn = ""
                        if mat_idx > -1 and mat_idx < len(row):
                            mat_cell = row[mat_idx]
                            if mat_cell:
                                match = re.search(r'\b(\d{9})\b', str(mat_cell))
                                if match:
                                    nsn = match.group(1)
                        
                        # Extract Quantity - Use OH QTY, fallback to AUTH QTY, fallback to 1
                        qty = 1
                        
                        # Try OH QTY first
                        if oh_qty_idx > -1 and oh_qty_idx < len(row):
                            qty_cell = row[oh_qty_idx]
                            if qty_cell:
                                try:
                                    qty_str = str(qty_cell).strip()
                                    # Extract first number found
                                    match = re.search(r'(\d+)', qty_str)
                                    if match:
                                        qty = int(match.group(1))
                                        print(f"DEBUG: OH QTY found: {qty}")
                                    else:
                                        print(f"DEBUG: OH QTY no number, defaulting to 1")
                                        qty = 1
                                except:
                                    qty = 1
                        else:
                            # Try AUTH QTY
                            if auth_qty_idx > -1 and auth_qty_idx < len(row):
                                qty_cell = row[auth_qty_idx]
                                if qty_cell:
                                    try:
                                        qty_str = str(qty_cell).strip()
                                        match = re.search(r'(\d+)', qty_str)
                                        if match:
                                            qty = int(match.group(1))
                                            print(f"DEBUG: AUTH QTY found: {qty}")
                                        else:
                                            print(f"DEBUG: AUTH QTY no number, defaulting to 1")
                                            qty = 1
                                    except:
                                        qty = 1
                            else:
                                # Default to 1
                                print(f"DEBUG: No QTY column found, defaulting to 1")
                                qty = 1
                        
                        items.append(BomItem(
                            line_no=len(items) + 1,
                            description=description[:100],
                            nsn=nsn,
                            qty=qty
                        ))
                        print(f"DEBUG: Added item {len(items)}")
    
    except Exception as e:
        print(f"CRITICAL ERROR in extraction: {e}")
        import traceback
        traceback.print_exc()
        return []
    
    print(f"DEBUG: Total items extracted: {len(items)}")
    return items


def generate_dd1750_from_pdf(bom_path: str, template_path: str, out_path: str, start_page: int = 0):
    """Generate DD1750 with robust file writing."""
    items = extract_items_from_pdf(bom_path, start_page)
    
    print(f"\nItems found: {len(items)}")
    
    if not items:
        # Fallback: Always create a file, even if empty
        reader = PdfReader(template_path)
        writer = PdfWriter()
        writer.add_page(reader.pages[0])
        output_pdf_path = out_path
        try:
            with open(output_pdf_path, 'wb') as f:
                writer.write(f)
                sys.stdout.flush()
            print(f"DEBUG: Wrote empty template to {output_pdf_path}")
        except:
            pass
        return output_pdf_path, 0
    
    total_pages = math.ceil(len(items) / ROWS_PER_PAGE)
    writer = PdfWriter()
    template = PdfReader(template_path)
    
    print(f"DEBUG: Creating {total_pages} pages")
    
    for page_num in range(total_pages):
        start_idx = page_num * ROWS_PER_PAGE
        end_idx = min((page_num + 1) * ROWS_PER_PAGE, len(items))
        page_items = items[start_idx:end_idx]
        
        print(f"DEBUG: Creating page {page_num} with {len(page_items)} items")
        
        # Create overlay
        packet = io.BytesIO()
        c = canvas.Canvas(packet, pagesize=(PAGE_W, PAGE_H))
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
        sys.stdout.flush()
        
        # Merge overlay
        overlay = PdfReader(packet)
        
        # Get template page (use first page for all)
        if page_num < len(template.pages):
            page = template.pages[page_num]
        else:
            page = template.pages[0]
        
        page.merge_page(overlay.pages[0])
        writer.add_page(page)
    
    # Write to file with verification
    output_pdf_path = out_path
    try:
        with open(output_pdf_path, 'wb') as f:
            writer.write(f)
        
        # Verify file was written
        sys.stdout.flush()
        
        if not os.path.exists(output_pdf_path):
            print(f"ERROR: Output file does not exist at {output_pdf_path}")
            # Try to write again as a last resort
            try:
                with open(output_pdf_path, 'wb') as f:
                    writer.write(f)
                    sys.stdout.flush()
            except:
                pass
        
        file_size = os.path.getsize(output_pdf_path)
        print(f"DEBUG: Output file size: {file_size} bytes")
        
        if file_size == 0:
            print("ERROR: Output file is 0 bytes - write failed")
            return output_pdf_path, 0
    
    except Exception as e:
        print(f"CRITICAL ERROR writing PDF: {e}")
        import traceback
        traceback.print_exc()
        
        # Fallback: Return template file if everything
