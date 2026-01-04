"""DD1750 core - Handle multi-page BOMs and text wrapping correctly."""

import io
import math
import re
from dataclasses import dataclass
from typing import List, Tuple, Dict, Any

import pdfplumber
from pypdf import PdfReader, PdfWriter
from reportlab.pdfgen import canvas
from reportlab.pdfbase import pdfmetrics


# Constants
ROWS_PER_PAGE = 18
PAGE_W, PAGE_H = 612.0, 792.0

# Column positions
X_BOX_L, X_BOX_R = 44.0, 88.0
X_CONTENT_L, X_CONTENT_R = 88.0, 365.0
X_UOI_L, X_UOI_R = 365.0, 408.5
X_INIT_L, X_INIT_R = 408.5, 453.5
X_SPARES_L, X_SPARES_R = 453.5, 514.5
X_TOTAL_L, X_TOTAL_R = 514.5, 566.0

Y_TABLE_TOP_LINE = 616.0
Y_TABLE_BOTTOM_LINE = 89.5
ROW_H = (Y_TABLE_TOP_LINE - Y_TABLE_BOTTOM_LINE) / ROWS_PER_PAGE
PAD_X = 3.0


@dataclass
class BomItem:
    line_no: int
    description: str
    nsn: str
    qty: int


def extract_items_from_pdf(pdf_path: str, start_page: int = 0) -> List[BomItem]:
    """Extract items handling text wrapping and multi-page BOMs."""
    items = []
    
    try:
        with pdfplumber.open(pdf_path) as pdf:
            # Process all pages
            for page_num, page in enumerate(pdf.pages[start_page:], start=start_page):
                print(f"Processing page {page_num}")
                
                # Extract tables if possible (most reliable)
                tables = page.extract_tables()
                
                if tables and len(tables) > 0:
                    print(f"  Found {len(tables)} tables")
                    
                    for table_idx, table in enumerate(tables):
                        if len(table) < 2:
                            continue
                        
                        # Find column indices
                        header = table[0]
                        col_indices = {}
                        
                        for idx, cell in enumerate(header):
                            if cell:
                                cell_text = str(cell).strip().upper()
                                if 'LV' in cell_text or 'LEVEL' in cell_text:
                                    col_indices['lv'] = idx
                                elif 'DESCRIPTION' in cell_text:
                                    col_indices['desc'] = idx
                                elif 'MATERIAL' in cell_text:
                                    col_indices['material'] = idx
                                elif 'QTY' in cell_text or 'QUANTITY' in cell_text:
                                    col_indices['qty'] = idx
                        
                        print(f"  Table {table_idx}: Found columns {col_indices}")
                        
                        # Process each row
                        current_item = None
                        for row_idx, row in enumerate(table[1:], start=1):
                            # Skip empty rows
                            if not any(cell for cell in row):
                                continue
                            
                            # Check if this is an LV = 'B' row
                            is_b_row = False
                            if 'lv' in col_indices and len(row) > col_indices['lv']:
                                lv_cell = row[col_indices['lv']]
                                if lv_cell and str(lv_cell).strip().upper() == 'B':
                                    is_b_row = True
                            
                            if is_b_row:
                                # Extract description
                                description = ""
                                if 'desc' in col_indices and len(row) > col_indices['desc']:
                                    desc_cell = row[col_indices['desc']]
                                    if desc_cell:
                                        description = str(desc_cell).strip()
                                        # Only take text before colon or parenthesis
                                        if ':' in description:
                                            description = description.split(':')[0].strip()
                                        if '(' in description:
                                            description = description.split('(')[0].strip()
                                        description = re.sub(r'\s+', ' ', description).strip()
                                
                                # Extract NSN from Material column
                                nsn = ""
                                if 'material' in col_indices and len(row) > col_indices['material']:
                                    material_cell = row[col_indices['material']]
                                    if material_cell:
                                        material_text = str(material_cell).strip()
                                        # Look for 9-digit NSN
                                        nsn_match = re.search(r'\b(\d{9})\b', material_text)
                                        if nsn_match:
                                            nsn = nsn_match.group(1)
                                
                                # Extract quantity
                                qty = 1
                                if 'qty' in col_indices and len(row) > col_indices['qty']:
                                    qty_cell = row[col_indices['qty']]
                                    if qty_cell:
                                        try:
                                            qty = int(str(qty_cell).strip())
                                        except:
                                            qty = 1
                                
                                if description:
                                    items.append(BomItem(
                                        line_no=len(items) + 1,
                                        description=description[:100],
                                        nsn=nsn,
                                        qty=qty
                                    ))
                                    print(f"    Added item {len(items)}: {description[:30]}... | NSN: {nsn}")
                
                # If no tables found or no items extracted, try text extraction
                if not items:
                    print("  No tables found, trying text extraction")
                    text = page.extract_text() or ""
                    lines = [ln.strip() for ln in text.split('\n') if ln.strip()]
                    
                    i = 0
                    while i < len(lines):
                        line = lines[i]
                        
                        # Look for patterns that indicate an item row
                        # Pattern 1: Starts with NSN then B then description
                        nsn_b_match = re.match(r'^(\d{9})\s+B\s+(.+)$', line)
                        if nsn_b_match:
                            nsn = nsn_b_match.group(1)
                            description = nsn_b_match.group(2)
                            
                            # Clean description
                            if ':' in description:
                                description = description.split(':')[0].strip()
                            if '(' in description:
                                description = description.split('(')[0].strip()
                            description = re.sub(r'\s+', ' ', description).strip()
                            
                            # Extract quantity
                            qty = 1
                            if description and description.split()[-1].isdigit():
                                qty = int(description.split()[-1])
                                description = ' '.join(description.split()[:-1])
                            
                            if description:
                                items.append(BomItem(
                                    line_no=len(items) + 1,
                                    description=description[:100],
                                    nsn=nsn,
                                    qty=qty
                                ))
                                print(f"    Added item {len(items)} (pattern 1): {description[:30]}...")
                        
                        # Pattern 2: Line with B and comma (description)
                        elif ' B ' in line and ',' in line:
                            parts = line.split()
                            if 'B' in parts:
                                b_index = parts.index('B')
                                
                                # Check for NSN before B
                                nsn = ""
                                if b_index > 0 and re.match(r'^\d{9}$', parts[b_index - 1]):
                                    nsn = parts[b_index - 1]
                                
                                # Get description after B
                                desc_parts = parts[b_index + 1:]
                                description = ' '.join(desc_parts)
                                
                                # Clean description
                                if ':' in description:
                                    description = description.split(':')[0].strip()
                                if '(' in description:
                                    description = description.split('(')[0].strip()
                                description = re.sub(r'\s+', ' ', description).strip()
                                
                                # Extract quantity
                                qty = 1
                                if description and description.split()[-1].isdigit():
                                    qty = int(description.split()[-1])
                                    description = ' '.join(description.split()[:-1])
                                
                                # If no NSN found, check nearby lines
                                if not nsn:
                                    for j in range(max(0, i-3), min(len(lines), i+3)):
                                        nsn_match = re.search(r'\b(\d{9})\b', lines[j])
                                        if nsn_match:
                                            nsn = nsn_match.group(1)
                                            break
                                
                                if description:
                                    items.append(BomItem(
                                        line_no=len(items) + 1,
                                        description=description[:100],
                                        nsn=nsn,
                                        qty=qty
                                    ))
                                    print(f"    Added item {len(items)} (pattern 2): {description[:30]}...")
                        
                        i += 1
    
    except Exception as e:
        print(f"ERROR in extraction: {e}")
        import traceback
        traceback.print_exc()
        return []
    
    print(f"=== Total items extracted: {len(items)} ===")
    return items


def generate_dd1750_from_pdf(bom_path, template_path, output_path, start_page=0):
    """Generate DD1750 with proper multi-page support."""
    try:
        items = extract_items_from_pdf(bom_path, start_page)
        
        print(f"\n=== FINAL ITEM LIST ({len(items)} items) ===")
        for i, item in enumerate(items, 1):
            print(f"{i:3d}. '{item.description[:40]}...' | NSN: {item.nsn} | Qty: {item.qty}")
        print("=== END LIST ===\n")
        
        if not items:
            # Return empty template
            reader = PdfReader(template_path)
            writer = PdfWriter()
            writer.add_page(reader.pages[0])
            with open(output_path, 'wb') as f:
                writer.write(f)
            return output_path, 0
        
        # Calculate how many pages we need
        total_pages = math.ceil(len(items) / ROWS_PER_PAGE)
        print(f"Creating {total_pages} pages for {len(items)} items")
        
        writer = PdfWriter()
        
        for page_num in range(total_pages):
            print(f"  Creating page {page_num + 1}")
            
            # Get items for this page
            start_idx = page_num * ROWS_PER_PAGE
            end_idx = min((page_num + 1) * ROWS_PER_PAGE, len(items))
            page_items = items[start_idx:end_idx]
            
            # Create overlay for this page
            packet = io.BytesIO()
            can = canvas.Canvas(packet, pagesize=(PAGE_W, PAGE_H))
            
            first_row_top = Y_TABLE_TOP_LINE - 5.0
            max_w = (X_CONTENT_R - X_CONTENT_L) - 2 * PAD_X
            
            for i, item in enumerate(page_items):
                y = first_row_top - (i * ROW_H)
                y_desc = y - 7.0
                y_nsn = y - 12.2
                
                # Box number (actual item number, not page position)
                can.setFont("Helvetica", 8)
                can.drawCentredString((X_BOX_L + X_BOX_R)/2, y_desc, str(item.line_no))
                
                # Description
                can.setFont("Helvetica", 7)
                desc = item.description
                if len(desc) > 50:
                    desc = desc[:47] + "..."
                can.drawString(X_CONTENT_L + PAD_X, y_desc, desc)
                
                # NSN
                if item.nsn:
                    can.setFont("Helvetica", 6)
                    nsn_text = f"NSN: {item.nsn}"
                    can.drawString(X_CONTENT_L + PAD_X, y_nsn, nsn_text)
                
                # Quantities
                can.setFont("Helvetica", 8)
                can.drawCentredString((X_UOI_L + X_UOI_R)/2, y_desc, "EA")
                can.drawCentredString((X_INIT_L + X_INIT_R)/2, y_desc, str(item.qty))
                can.drawCentredString((X_SPARES_L + X_SPARES_R)/2, y_desc, "0")
                can.drawCentredString((X_TOTAL_L + X_TOTAL_R)/2, y_desc, str(item.qty))
            
            can.save()
            packet.seek(0)
            
            # Merge with template
            overlay = PdfReader(packet)
            template_page = PdfReader(template_path).pages[0]
            template_page.merge_page(overlay.pages[0])
            writer.add_page(template_page)
        
        with open(output_path, 'wb') as f:
            writer.write(f)
        
        print(f"Successfully created {output_path} with {len(items)} items on {total_pages} pages")
        return output_path, len(items)
        
    except Exception as e:
        print(f"CRITICAL ERROR: {e}")
        import traceback
        traceback.print_exc()
        # Return empty template on any error
        reader = PdfReader(template_path)
        writer = PdfWriter()
        writer.add_page(reader.pages[0])
        with open(output_path, 'wb') as f:
            writer.write(f)
        return output_path, 0
