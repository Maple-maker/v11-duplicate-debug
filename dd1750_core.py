"""DD1750 core - Fixed with proper positioning and quantity extraction."""

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


# Use letter size
PAGE_W, PAGE_H = letter

# Column positions
X_BOX_L, X_BOX_R = 44.0, 88.0
X_CONTENT_L, X_CONTENT_R = 88.0, 365.0
X_UOI_L, X_UOI_R = 365.0, 408.5
X_INIT_L, X_INIT_R = 408.5, 453.5
X_SPARES_L, X_SPARES_R = 453.5, 514.5
X_TOTAL_L, X_TOTAL_R = 514.5, 566.0

# Table row positions
Y_TABLE_TOP_LINE = 616.0
Y_TABLE_BOTTOM_LINE = 89.5
ROWS_PER_PAGE = 18
ROW_H = (Y_TABLE_TOP_LINE - Y_TABLE_BOTTOM_LINE) / ROWS_PER_PAGE
PAD_X = 3.0

# Admin field positions (adjusted for the template in screenshot)
ADMIN_Y = {
    'unit': 720,           # Unit box (top left)
    'requisition': 720,    # Requisition box (top middle-right)
    'page': 720,           # Page box (top right)
    'date': 695,           # Date box (second row left)
    'order': 695,          # Order box (second row middle)
    'boxes': 695,          # Boxes box (second row right)
    'packed_by': 115,      # Packed by (bottom section)
    'received_by': 115,    # Received by (bottom section)
}


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
                    
                    # Print header to debug
                    print(f"\nDEBUG: Header row: {[str(c) if c else 'EMPTY' for c in header]}")
                    
                    for i, cell in enumerate(header):
                        if cell:
                            text = str(cell).upper()
                            print(f"DEBUG: Column {i}: '{text}'")
                            if 'LV' in text or 'LEVEL' in text:
                                lv_idx = i
                                print(f"  -> Found LV column at index {i}")
                            elif 'DESC' in text or 'NOMENCLATURE' in text:
                                desc_idx = i
                                print(f"  -> Found DESC column at index {i}")
                            elif 'MATERIAL' in text or 'NSN' in text:
                                mat_idx = i
                                print(f"  -> Found MATERIAL column at index {i}")
                            elif 'AUTH' in text or 'QTY' in text or 'QUANTITY' in text:
                                qty_idx = i
                                print(f"  -> Found QTY column at index {i}")
                    
                    print(f"\nDEBUG: Using columns - LV:{lv_idx}, DESC:{desc_idx}, MAT:{mat_idx}, QTY:{qty_idx}")
                    
                    if lv_idx is None or desc_idx is None:
                        print("SKIPPING TABLE: Missing LV or DESC column")
                        continue
                    
                    row_num = 0
                    for row in table[1:]:
                        row_num += 1
                        if not any(cell for cell in row if cell):
                            continue
                        
                        lv_cell = row[lv_idx] if lv_idx < len(row) else None
                        if not lv_cell:
                            continue
                        
                        lv_text = str(lv_cell).strip().upper()
                        if lv_text != 'B':
                            continue
                        
                        # Get description
                        desc_cell = row[desc_idx] if desc_idx < len(row) else None
                        description = ""
                        if desc_cell:
                            desc_text = str(desc_cell).strip()
                            print(f"\nDEBUG Row {row_num}: Raw description:\n  '{desc_text}'")
                            
                            lines = desc_text.split('\n')
                            description = lines[1].strip() if len(lines) >= 2 else lines[0].strip()
                            
                            # Clean description
                            if '(' in description:
                                description = description.split('(')[0].strip()
                            description = re.sub(r'\s+(WTY|ARC|CIIC|UI|SCMC|EA|AY|9K|9G)$', '', description, flags=re.IGNORECASE)
                            description = re.sub(r'\s+', ' ', description).strip()
                            
                            print(f"  -> Cleaned: '{description}'")
                        
                        if not description:
                            continue
                        
                        # Get NSN
                        nsn = ""
                        if mat_idx is not None and mat_idx < len(row):
                            mat_cell = row[mat_idx]
                            if mat_cell:
                                mat_text = str(mat_cell).strip()
                                print(f"  -> Material cell: '{mat_text}'")
                                match = re.search(r'\b(\d{9})\b', mat_text)
                                if match:
                                    nsn = match.group(1)
                                    print(f"  -> Found NSN: {nsn}")
                        
                        # Get quantity - try multiple strategies
                        qty = 1
                        if qty_idx is not None and qty_idx < len(row):
                            qty_cell = row[qty_idx]
                            if qty_cell:
                                qty_text = str(qty_cell).strip()
                                print(f"  -> Qty cell raw: '{qty_text}'")
                                
                                # Try to extract number
                                qty_match = re.search(r'(\d+)', qty_text)
                                if qty_match:
                                    qty = int(qty_match.group(1))
                                    print(f"  -> Extracted qty: {qty}")
                                else:
                                    print(f"  -> Could not extract qty from '{qty_text}', defaulting to 1")
                        else:
                            print(f"  -> No qty column found, defaulting to 1")
                        
                        items.append(BomItem(len(items) + 1, description[:100], nsn, qty))
                        print(f"  -> ADDED ITEM #{len(items)}: {description[:30]}... | NSN:{nsn} | Qty:{qty}")
    
    except Exception as e:
        print(f"ERROR: {e}")
        import traceback
        traceback.print_exc()
    
    print(f"\n=== TOTAL ITEMS EXTRACTED: {len(items)} ===\n")
    return items


def generate_dd1750_from_pdf(bom_path, template_path, output_path, start_page=0, admin_data=None):
    if admin_data is None:
        admin_data = {}
    
    print(f"\n{'='*60}")
    print(f"ADMIN DATA RECEIVED: {admin_data}")
    print(f"{'='*60}\n")
    
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
    print(f"\nCreating {total_pages} pages for {len(items)} items\n")
    
    writer = PdfWriter()
    
    for page_num in range(total_pages):
        try:
            start_idx = page_num * ROWS_PER_PAGE
            end_idx = min((page_num + 1) * ROWS_PER_PAGE, len(items))
            page_items = items[start_idx:end_idx]
            
            # Create overlay
            packet = io.BytesIO()
            c = canvas.Canvas(packet, pagesize=letter)
            
            first_row_top = Y_TABLE_TOP_LINE - 5.0
            
            # Draw items
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
            
            # Draw admin fields at TOP of page (not in middle!)
            print(f"\nDrawing admin fields at top of page...")
            
            # UNIT (top left)
            if admin_data.get('unit'):
                c.setFont("Helvetica", 8)
                c.drawString(50, ADMIN_Y['unit'], admin_data['unit'][:25])
                print(f"  -> Drew Unit at y={ADMIN_Y['unit']}: '{admin_data['unit']}'")
            
            # REQUISITION NO.
            if admin_data.get('requisition_no'):
                c.setFont("Helvetica", 8)
                c.drawString(250, ADMIN_Y['requisition'], f"REQ: {admin_data['requisition_no']}")
                print(f"  -> Drew Requisition at y={ADMIN_Y['requisition']}")
            
            # PAGE
            if total_pages > 1:
                c.setFont("Helvetica", 8)
                c.drawString(480, ADMIN_Y['page'], f"PAGE {page_num + 1}/{total_pages}")
                print(f"  -> Drew Page at y={ADMIN_Y['page']}")
            
            # DATE
            if admin_data.get('date'):
                c.setFont("Helvetica", 8)
                c.drawString(50, ADMIN_Y['date'], admin_data['date'])
                print(f"  -> Drew Date at y={ADMIN_Y['date']}")
            
            # ORDER NO.
            if admin_data.get('order_no'):
                c.setFont("Helvetica", 8)
                c.drawString(250, ADMIN_Y['order'], f"ORDER: {admin_data['order_no']}")
                print(f"  -> Drew Order at y={ADMIN_Y['order']}")
            
            # TOTAL NO. OF BOXES
            if admin_data.get('num_boxes'):
                c.setFont("Helvetica", 8)
                c.drawString(450, ADMIN_Y['boxes'], f"BOXES: {admin_data['num_boxes']}")
                print(f"  -> Drew Boxes at y={ADMIN_Y['boxes']}")
            
            # PACKED BY (bottom section - on EVERY page)
            if admin_data.get('packed_by'):
                c.setFont("Helvetica", 8)
                c.drawString(44, ADMIN_Y['packed_by'], f"PACKED BY: {admin_data['packed_by']}")
                c.setFont("Helvetica", 6)
                c.drawString(44, ADMIN_Y['packed_by'] - 10, "(Signature)")
                print(f"  -> Drew Packed By at y={ADMIN_Y['packed_by']}")
            
            # RECEIVED BY (bottom section - on EVERY page)
            c.setFont("Helvetica", 8)
            c.drawString(300, ADMIN_Y['received_by'], "RECEIVED BY:")
            c.setFont("Helvetica", 6)
            c.drawString(300, ADMIN_Y['received_by'] - 10, "(Signature)")
            print(f"  -> Drew Received By at y={ADMIN_Y['received_by']}")
            
            c.save()
            packet.seek(0)
            
            # Merge
            overlay = PdfReader(packet)
            template = PdfReader(template_path)
            page = template.pages[0]
            page.merge_page(overlay.pages[0])
            writer.add_page(page)
            
            print(f"Page {page_num + 1} created successfully\n")
            
        except Exception as e:
            print(f"ERROR on page {page_num + 1}: {e}")
            import traceback
            traceback.print_exc()
            try:
                template = PdfReader(template_path)
                writer.add_page(template.pages[0])
            except:
                pass
    
    # Write output
    try:
        with open(output_path, 'wb') as f:
            writer.write(f)
        print(f"\nSUCCESS: Wrote {output_path}")
    except Exception as e:
        print(f"ERROR writing output: {e}")
    
    return output_path, len(items)
