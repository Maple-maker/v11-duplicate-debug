"""DD1750 core - Fixed version with correct NSN and description extraction."""

import io
import math
import re
from dataclasses import dataclass
from typing import List, Tuple, Dict

import pdfplumber
from pypdf import PdfReader, PdfWriter
from reportlab.pdfgen import canvas
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont


# Register fonts for better rendering
try:
    pdfmetrics.registerFont(TTFont('Arial', 'arial.ttf'))
    DEFAULT_FONT = 'Arial'
except:
    DEFAULT_FONT = 'Helvetica'


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


def _identify_columns(header: List) -> Dict[str, int]:
    """Identify column positions based on header content."""
    indices = {}
    
    for idx, cell in enumerate(header):
        if not cell:
            continue
        
        cell_text = str(cell).strip().upper()
        
        # LV/Level column
        if 'LV' in cell_text or 'LEVEL' in cell_text or 'L/V' in cell_text:
            indices['lv'] = idx
        
        # Description column
        elif 'DESCRIPTION' in cell_text or 'DESC' in cell_text:
            indices['desc'] = idx
        
        # Material/NSN column
        elif 'MATERIAL' in cell_text:
            indices['material'] = idx
        
        # Quantity column
        elif 'QTY' in cell_text or 'AUTH' in cell_text:
            indices['qty'] = idx
    
    return indices


def _extract_middle_text(text: str) -> str:
    """Extract the middle text from description (not top, not bottom)."""
    if not text:
        return ""
    
    # Split by newlines
    lines = [ln.strip() for ln in text.split('\n') if ln.strip()]
    
    if not lines:
        return ""
    
    # If we have multiple lines, take the SECOND line (middle text)
    # This is the text that's horizontally aligned with "B" in LV column
    if len(lines) >= 2:
        middle_line = lines[1]  # Second line is the middle text
    else:
        middle_line = lines[0]
    
    # Clean up the middle line
    # Remove parenthetical text (anything after opening parenthesis)
    if '(' in middle_line:
        middle_line = middle_line.split('(')[0].strip()
    
    # Remove trailing codes (WTY, ARC, CIIC, UI, SCMC, EA, AY, etc.)
    middle_line = re.sub(r'\s+(WTY|ARC|CIIC|UI|SCMC|EA|AY|9K|9G)$', '', middle_line, flags=re.IGNORECASE)
    
    # Clean up whitespace
    middle_line = re.sub(r'\s+', ' ', middle_line).strip()
    
    # Remove leading/trailing punctuation
    middle_line = middle_line.strip('.,-')
    
    return middle_line[:100]


def _extract_nsn(text: str) -> str:
    """Extract NSN (9-digit number) from Material column."""
    if not text:
        return ""
    
    # Look for 9-digit NSN pattern
    nsn_match = re.search(r'\b(\d{9})\b', text)
    if nsn_match:
        return nsn_match.group(1)
    
    return ""


def _extract_quantity(text: str) -> int:
    """Extract quantity from Qty column."""
    if not text:
        return 1
    
    try:
        return int(str(text).strip())
    except:
        return 1


def extract_items_from_pdf(pdf_path: str, start_page: int = 0) -> List[BomItem]:
    """Extract items from BOM PDF - table-based extraction only."""
    items = []
    
    try:
        with pdfplumber.open(pdf_path) as pdf:
            for page_num, page in enumerate(pdf.pages[start_page:], start=start_page):
                print(f"\n=== Processing page {page_num} ===")
                
                # Extract tables
                tables = page.extract_tables()
                
                if not tables:
                    print("  No tables found on this page")
                    continue
                
                print(f"  Found {len(tables)} tables")
                
                for table_idx, table in enumerate(tables):
                    if len(table) < 2:
                        print(f"  Table {table_idx} is empty, skipping")
                        continue
                    
                    # Find column indices from header row
                    header = table[0]
                    col_indices = _identify_columns(header)
                    
                    print(f"  Table {table_idx}: Columns = {col_indices}")
                    
                    if 'lv' not in col_indices or 'desc' not in col_indices:
                        print(f"  Table {table_idx}: Missing LV or DESC column, skipping")
                        continue
                    
                    # Process each data row (skip header)
                    for row_idx, row in enumerate(table[1:], start=1):
                        # Skip empty rows
                        if not any(cell for cell in row if cell):
                            continue
                        
                        # Get LV value
                        lv_idx = col_indices['lv']
                        lv_cell = row[lv_idx] if lv_idx < len(row) else None
                        
                        if not lv_cell:
                            continue
                        
                        # Only process rows where LV = 'B'
                        lv_text = str(lv_cell).strip().upper()
                        if lv_text != 'B':
                            continue
                        
                        print(f"    Row {row_idx}: Found LV='B'")
                        
                        # Get description from DESC column
                        desc_idx = col_indices['desc']
                        desc_cell = row[desc_idx] if desc_idx < len(row) else None
                        
                        if not desc_cell:
                            print(f"    Row {row_idx}: No description, skipping")
                            continue
                        
                        # Extract middle text (second line)
                        description = _extract_middle_text(str(desc_cell))
                        
                        if not description:
                            print(f"    Row {row_idx}: Empty description after cleaning, skipping")
                            continue
                        
                        # Get NSN from Material column (same row)
                        nsn = ""
                        if 'material' in col_indices:
                            mat_idx = col_indices['material']
                            mat_cell = row[mat_idx] if mat_idx < len(row) else None
                            if mat_cell:
                                nsn = _extract_nsn(str(mat_cell))
                        
                        print(f"    Row {row_idx}: NSN from Material = '{nsn}'")
                        
                        # Get quantity from Qty column (same row)
                        qty = 1
                        if 'qty' in col_indices:
                            qty_idx = col_indices['qty']
                            qty_cell = row[qty_idx] if qty_idx < len(row) else None
                            if qty_cell:
                                qty = _extract_quantity(str(qty_cell))
                        
                        # Add the item
                        items.append(BomItem(
                            line_no=len(items) + 1description,
                            n,
                            description=sn=nsn,
                            qty=qty
                        ))
                        
                        print(f"    Added item {len(items)}: '{description[:50]}' | NSN: {nsn} | Qty: {qty}")
    
    except Exception as e:
        print(f"ERROR: {e}")
        import traceback
        traceback.print_exc()
        return []
    
    print(f"\n=== Total items extracted: {len(items)} ===")
    return items


def generate_dd1750_from_pdf(bom_path, template_path, output_path, start_page=0):
    """Generate DD1750 PDF."""
    try:
        items = extract_items_from_pdf(bom_path, start_page)
        
        print(f"\n=== FINAL ITEM LIST ({len(items)} items) ===")
        for i, item in enumerate(items, 1):
            print(f"{i:3d}. '{item.description}' | NSN: {item.nsn} | Qty: {item.qty}")
        print("=== END LIST ===\n")
        
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
        print(f"Creating {total_pages} pages for {len(items)} items")
        
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
                can.setFont(DEFAULT_FONT, 8)
                can.drawCentredString((X_BOX_L + X_BOX_R)/2, y_desc, str(item.line_no))
                
                # Description
                can.setFont(DEFAULT_FONT, 7)
                desc = item.description
                if len(desc) > 50:
                    desc = desc[:47] + "..."
                can.drawString(X_CONTENT_L + PAD_X, y_desc, desc)
                
                # NSN
                if item.nsn:
                    can.setFont(DEFAULT_FONT, 6)
                    can.drawString(X_CONTENT_L + PAD_X, y_nsn, f"NSN: {item.nsn}")
                
                # Quantities
                can.setFont(DEFAULT_FONT, 8)
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
        import traceback
        traceback.print_exc()
        # Return empty template
        reader
