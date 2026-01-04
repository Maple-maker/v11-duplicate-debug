"""DD1750 core - Simple extraction of middle text from description boxes."""

import io
import math
import re
from dataclasses import dataclass
from typing import List, Tuple

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


def extract_middle_text(description: str) -> str:
    """Extract the middle text from a description box."""
    if not description:
        return ""
    
    # Split by newlines
    lines = [ln.strip() for ln in description.split('\n') if ln.strip()]
    
    if len(lines) == 0:
        return ""
    
    # Strategy 1: If we have multiple lines, take the middle one
    if len(lines) >= 2:
        # Usually the middle text is line 1 (0-indexed)
        middle_line = lines[1] if len(lines) > 1 else lines[0]
        
        # Clean it up
        middle_line = re.sub(r'$$.*?$$', '', middle_line)  # Remove parentheses
        middle_line = re.sub(r'\s+', ' ', middle_line).strip()
        
        # If it's still good, return it
        if middle_line and len(middle_line) > 2:
            return middle_line[:100]
    
    # Strategy 2: If single line, try to extract the specific part
    # Look for patterns like "7FT CHAIN" or specific measurements
    single_line = lines[0]
    
    # Remove parenthetical text
    single_line = re.sub(r'$$.*?$$', '', single_line)
    
    # Look for measurement patterns (7FT, 14FT, etc.)
    measurement_match = re.search(r'(\d+\s*(FT|IN|MM|CM|M)\s+\w+)', single_line, re.IGNORECASE)
    if measurement_match:
        return measurement_match.group(1).strip()[:100]
    
    # Look for specific patterns after comma
    if ',' in single_line:
        parts = single_line.split(',')
        if len(parts) >= 2:
            # Take the part after the first comma
            after_comma = parts[1].strip()
            if after_comma and len(after_comma) > 2:
                return after_comma[:100]
    
    # Fallback: return the whole line, cleaned up
    single_line = re.sub(r'\s+', ' ', single_line).strip()
    return single_line[:100]


def extract_items_from_pdf(pdf_path: str, start_page: int = 0) -> List[BomItem]:
    """Simple, reliable extraction."""
    items = []
    
    try:
        with pdfplumber.open(pdf_path) as pdf:
            for page_num, page in enumerate(pdf.pages[start_page:], start=start_page):
                print(f"Processing page {page_num}")
                
                # Try table extraction first (most reliable)
                tables = page.extract_tables()
                
                if tables:
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
                        
                        print(f"  Table {table_idx}: Columns found: {list(col_indices.keys())}")
                        
                        # Process rows
                        for row_idx, row in enumerate(table[1:], start=1):
                            # Skip empty rows
                            if not any(cell for cell in row):
                                continue
                            
                            # Check if LV = 'B'
                            if 'lv' in col_indices and len(row) > col_indices['lv']:
                                lv_cell = row[col_indices['lv']]
                                if lv_cell and str(lv_cell).strip().upper() == 'B':
                                    # Extract description
                                    description = ""
                                    if 'desc' in col_indices and len(row) > col_indices['desc']:
                                        desc_cell = row[col_indices['desc']]
                                        if desc_cell:
                                            description = str(desc_cell).strip()
                                    
                                    # Extract middle text
                                    middle_text = extract_middle_text(description)
                                    
                                    # Extract NSN
                                    nsn = ""
                                    if 'material' in col_indices and len(row) > col_indices['material']:
                                        material_cell = row[col_indices['material']]
                                        if material_cell:
                                            material_text = str(material_cell).strip()
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
                                    
                                    if middle_text:
                                        items.append(BomItem(
                                            line_no=len(items) + 1,
                                            description=middle_text,
                                            nsn=nsn,
                                            qty=qty
                                        ))
                                        print(f"    Item {len(items)}: '{middle_text[:40]}...' | NSN: {nsn}")
                
                # Fallback: Text extraction if no tables
                if not items:
                    print("  No tables found, using text extraction")
                    text = page.extract_text() or ""
                    lines = [ln.strip() for ln in text.split('\n') if ln.strip()]
                    
                    i = 0
                    while i < len(lines):
                        line = lines[i]
                        
                        # Look for LV = 'B' pattern
                        if re.match(r'^\d{9}\s+B\s+', line) or (' B ' in line and ',' in line):
                            # Extract NSN if present at start
                            nsn = ""
                            nsn_match = re.match(r'^(\d{9})\s+', line)
                            if nsn_match:
                                nsn = nsn_match.group(1)
                            
                            # Get the full description text
                            full_desc = line
                            
                            # Look ahead for continuation lines
                            j = i + 1
                            while j < len(lines) and j < i + 3:
                                next_line = lines[j]
                                # If next line doesn't start a new item, add it to description
                                if not re.match(r'^\d{9}\s+B\s+', next_line) and not (' B ' in next_line):
                                    full_desc += " " + next_line
                                    j += 1
                                else:
                                    break
                            
                            # Extract middle text
                            middle_text = extract_middle_text(full_desc)
                            
                            # If no NSN found, check nearby lines
                            if not nsn:
                                for k in range(max(0, i-2), min(len(lines), i+3)):
                                    nsn_match = re.search(r'\b(\d{9})\b', lines[k])
                                    if nsn_match:
                                        nsn = nsn_match.group(1)
                                        break
                            
                            # Extract quantity (look for number at end)
                            qty = 1
                            if middle_text and middle_text.split()[-1].isdigit():
                                qty = int(middle_text.split()[-1])
                                middle_text = ' '.join(middle_text.split()[:-1])
                            
                            if middle_text:
                                items.append(BomItem(
                                    line_no=len(items) + 1,
                                    description=middle_text[:100],
                                    nsn=nsn,
                                    qty=qty
                                ))
                                print(f"    Item {len(items)} (text): '{middle_text[:40]}...'")
                        
                        i += 1
    
    except Exception as e:
        print(f"ERROR: {e}")
        return []
    
    print(f"\n=== Total items: {len(items)} ===")
    return items


def generate_dd1750_from_pdf(bom_path, template_path, output_path, start_page=0):
    """Generate DD1750 - Simple and reliable."""
    try:
        items = extract_items_from_pdf(bom_path, start_page)
        
        print(f"\n=== FINAL ITEMS ({len(items)} total) ===")
        for i, item in enumerate(items, 1):
            print(f"{i:3d}. '{item.description}' | NSN: {item.nsn} | Qty: {item.qty}")
        
        if not items:
            # Return empty template
            reader = PdfReader(template_path)
            writer = PdfWriter()
            writer.add_page(reader.pages[0])
            with open(output_path, 'wb') as f:
                writer.write(f)
            return output_path, 0
        
        # Calculate pages needed
        total_pages = math.ceil(len(items) / ROWS_PER_PAGE)
        print(f"\nCreating {total_pages} pages")
        
        writer = PdfWriter()
        
        for page_num in range(total_pages):
            # Get items for this page
            start_idx = page_num * ROWS_PER_PAGE
            end_idx = min((page_num + 1) * ROWS_PER_PAGE, len(items))
            page_items = items[start_idx:end_idx]
            
            # Create overlay
            packet = io.BytesIO()
            can = canvas.Canvas(packet, pagesize=(PAGE_W, PAGE_H))
            
            first_row_top = Y_TABLE_TOP_LINE - 5.0
            
            for i, item in enumerate(page_items):
                y = first_row_top - (i * ROW_H)
                y_desc = y - 7.0
                y_nsn = y - 12.2
                
                # Box number
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
                    can.drawString(X_CONTENT_L + PAD_X, y_nsn, f"NSN: {item.nsn}")
                
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
        
        return output_path, len(items)
        
    except Exception as e:
        print(f"CRITICAL ERROR: {e}")
        # Return empty template
        reader = PdfReader(template_path)
        writer = PdfWriter()
        writer.add_page(reader.pages[0])
        with open(output_path, 'wb') as f:
            writer.write(f)
        return output_path, 0
