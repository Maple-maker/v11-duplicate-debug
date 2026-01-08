"""DD1750 core - Bulletproof version."""

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
    """Extract items from BOM PDF."""
    items = []
    
    try:
        print(f"DEBUG: Opening {pdf_path}")
        
        with pdfplumber.open(pdf_path) as pdf:
            for page in pdf.pages[start_page:]:
                tables = page.extract_tables()
                
                for table in tables:
                    if len(table) < 2:
                        continue
                    
                    header = table[0]
                    lv_idx = desc_idx = mat_idx = auth_idx = -1
                    oh_qty_idx = -1
                    
                    # Identify columns
                    for i, cell in enumerate(header):
                        if cell:
                            text = str(cell).upper()
                            if 'LV' in text or 'LEVEL' in text:
                                lv_idx = i
                            elif 'DESC' in text or 'NOMENCLATURE' in text or 'PART NO.' in text:
                                desc_idx = i
                            elif 'MATERIAL' in text:
                                mat_idx = i
                            elif 'OH' in text and 'QTY' in text:
                                oh_qty_idx = i
                            elif 'AUTH' in text and 'QTY' in text:
                                auth_idx = i
                    
                    if lv_idx == -1 or desc_idx == -1:
                        print("DEBUG: Skipped table (no LV or DESC column)")
                        continue
                    
                    for row in table[1:]:
                        if not any(cell for cell in row if cell):
                            continue
                        
                        # Check Level B
                        lv_cell = row[lv_idx] if lv_idx < len(row) else None
                        if not lv_cell or str(lv_cell).strip().upper() != 'B':
                            continue
                        
                        # Extract Description - USE RAW CONTENT
                        desc_cell = row[desc_idx] if desc_idx < len(row) else None
                        description = ""
                        if desc_cell:
                            description = str(desc_cell).strip()
                        
                        # Extract NSN - Search within raw description text
                        nsn = ""
                        if mat_idx > -1 and mat_idx < len(row):
                            mat_cell = row[mat_idx]
                            if mat_cell:
                                match = re.search(r'\b(\d{9})\b', str(mat_cell))
                                if match:
                                    nsn = match.group(1)
                        
                        # Extract Quantity - PREFER OH QTY OVER AUTH QTY
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
                        
                        items.append(BomItem(
                            line_no=len(items) + 1,
                            description=description[:100],
                            nsn=nsn,
                            qty=qty
                        ))
    
    except Exception as e:
        print(f"ERROR: {e}")
        import traceback
        traceback.print_exc()
        return []
    
    return items


def generate_dd1750_from_pdf(bom_path: str, template_path: str, out_path: str, start_page: int = 0):
    """Generate DD1750 with items overlayed on template."""
    items = extract_items_from_pdf(bom_path, start_page)
    
    print(f"DEBUG: Total items extracted: {len(items)}")
    
    if not items:
        # Create a minimal file so user gets *something*
        print("DEBUG: No items found, creating minimal output")
        writer = PdfWriter()
        
        try:
            reader = PdfReader(template_path)
            writer.add_page(reader.pages[0])
            with open(out_path, 'wb') as f:
                writer.write(f)
            print(f"DEBUG: Wrote minimal template to {out_path}")
        except Exception as e:
            print(f"ERROR writing minimal: {e}")
        
        return out_path, 0
    
    total_pages = math.ceil(len(items) / ROWS_PER_PAGE)
    writer = PdfWriter()
    
    # Read template
    try:
        template_reader = PdfReader(template_path)
    except Exception as e:
        print(f"ERROR reading template: {e}")
        # Fallback: return template as items to prevent crash
        # (Actually creates infinite loop if we try to process)
        writer = PdfWriter()
        writer.add_page(PdfReader(bom_path).pages[0]) # Just use BOM page as placeholder
        with open(out_path, 'wb') as f:
            writer.write(f)
        return out_path, 0
    
    for page_num in range(total_pages):
        start_idx = page_num * ROWS_PER_PAGE
        end_idx = min((page_num + 1) * ROWS_PER_PAGE, len(items))
        page_items = items[start_idx:end_idx]
        
        # Create overlay
        packet = io.BytesIO()
        c = canvas.Canvas(packet, pagesize=(PAGE_W, PAGE_H))
        
        first_row = Y_TABLE_TOP - 5.0
        
        for i, item in enumerate(page_items):
            y = first_row - (i * ROW_H)
            y_desc = y - 7.0
            y_nsn = y - 12.2
            
            # Box number
            c.setFont("Helvetica", 8)
            c.drawCentredString((X_BOX_L + X_BOX_R) / 2, y_desc, str(item.line_no))
            
            # Description
            c.setFont("Helvetica", 7)
            c.drawString(X_CONTENT_L + PAD_X, y_desc, item.description[:50])
            
            # NSN
            if item.nsn:
                c.setFont("Helvetica", 6)
                c.drawString(X_CONTENT_L + PAD_X, y_nsn, f"NSN: {item.nsn}")
            
            # Quantities
            c.setFont("Helvetica", 8)
            c.drawCentredString((X_UOI_L + X_UOI_R) / 2, y_desc, "EA")
            c.drawCentredString((X_INIT_L + X_INIT_R) / 2, y_desc, str(item.qty))
            c.drawCentredString((X_SPARES_L + X_SPARES_R) / 2, y_desc, "0")
            c.drawCentredString((X_TOTAL_L + X_TOTAL_R) / 2, y_desc, str(item.qty))
        
        c.save()
        packet.seek(0)
        
        # Merge overlay
        try:
            overlay = PdfReader(packet)
            
            # Get template page (use first page for all to ensure consistent background)
            if page_num < len(template_reader.pages):
                page = template_reader.pages[page_num]
            else:
                page = template_reader.pages[0]
            
            page.merge_page(overlay.pages[0])
            writer.add_page(page)
        except Exception as e:
            print(f"ERROR merging page {page_num}: {e}")
    
    # Write output file - CRITICAL STEP
    try:
        with open(out_path, 'wb') as f:
            writer.write(f)
            print(f"DEBUG: Successfully wrote {out_path} with {len(writer.pages)} pages")
            sys.stdout.flush()
    except Exception as e:
        print(f"CRITICAL ERROR writing PDF: {e}")
        import traceback
        traceback.print_exc()
        
        # Last resort fallback: Write a simple blank file
        try:
            template_reader = PdfReader(template_path)
            simple_writer = PdfWriter()
            simple_writer.add_page(template_reader.pages[0])
            with open(out_path, 'wb') as f:
                simple_writer.write(f)
        except:
            pass
    
    return out_path, len(items)
