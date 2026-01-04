"""DD1750 core - Simple, reliable table-based extraction."""

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


def extract_items_from_pdf(pdf_path: str, start_page: int = 0) -> List[BomItem]:
    """Simple, reliable table-based extraction."""
    items = []
    
    try:
        with pdfplumber.open(pdf_path) as pdf:
            for page_num, page in enumerate(pdf.pages[start_page:], start=start_page):
                print(f"\n=== Processing page {page_num} ===")
                
                # Try table extraction ONLY
                tables = page.extract_tables()
                
                if not tables:
                    print("  No tables found on this page")
                    continue
                
                print(f"  Found {len(tables)} tables")
                
                for table_idx, table in enumerate(tables):
                    if len(table) < 2:
                        print(f"  Table {table_idx} has no data rows, skipping")
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
                            elif 'QTY' in cell_text or 'AUTH QTY' in cell_text:
                                col_indices['qty'] = idx
                    
                    print(f"  Table {table_idx}: Columns found: LV={col_indices.get('lv')}, DESC={col_indices.get('desc')}, MAT={col_indices.get('material')}, QTY={col_indices.get('qty')}")
                    
                    # Process each data row (skip header)
                    for row_idx, row in enumerate(table[1:], start=1):
                        # Skip empty rows
                        if not any(cell for cell in row if cell):
                            continue
                        
                        # Check if LV = 'B'
                        lv_idx = col_indices.get('lv')
                        if lv_idx is None or lv_idx >= len(row):
                            continue
                        
                        lv_cell = row[lv_idx]
                        if not lv_cell:
                            continue
                        
                        lv_text = str(lv_cell).strip().upper()
                        if lv_text != 'B':
                            continue
                        
                        print(f"    Row {row_idx}: Found LV='B'")
                        
                        # Get description from DESC column
                        description = ""
                        desc_idx = col_indices.get('desc')
                        if desc_idx is not None and desc_idx < len(row):
                            desc_cell = row[desc_idx]
                            if desc_cell:
                                desc_text = str(desc_cell).strip()
                                
                                # Split by newlines - get the SECOND line (middle text)
                                lines = desc_text.split('\n')
                                
                                if len(lines) >= 2:
                                    # Take line 1 (second line, which is the middle text)
                                    description = lines[1].strip()
                                elif len(lines) == 1:
                                    # Only one line - clean it up
                                    description = lines[0].strip()
                                
                                # Clean description
                                # Remove parenthetical text
                                if '(' in description:
                                    description = description.split('(')[0].strip()
                                # Remove trailing codes
                                description = re.sub(r'\s+(WTY|ARC|CIIC|UI|SCMC|EA|AY|9K|9G)\s*$', '', description, flags=re.IGNORECASE)
                                description = re.sub(r'\s+', ' ', description).strip()
                        
                        if not description:
                            print(f"    Row {row_idx}: No description found, skipping")
                            continue
                        
                        # Get NSN from Material column
                        nsn = ""
                        mat_idx = col_indices.get('material')
                        if mat_idx is not None and mat_idx < len(row):
                            mat_cell = row[mat_idx]
                            if mat_cell:
                                mat_text = str(mat_cell).strip()
                                # Look for 9-digit NSN
                                nsn_match = re.search(r'\b(\d{9})\b', mat_text)
                                if nsn_match:
                                    nsn = nsn_match.group(1)
                        
                        # Get quantity from QTY column
                        qty = 1
                        qty_idx = col_indices.get('qty')
                        if qty_idx is not None and qty_idx < len(row):
                            qty_cell = row[qty_idx]
                            if qty_cell:
                                try:
                                    qty = int(str(qty_cell).strip())
                                except:
                                    qty = 1
                        
                        # Add the item
                        items.append(BomItem(
                            line_no=len(items) + 1,
                            description=description,
                            nsn=nsn,
                            qty=qty
                        ))
                        
                        print(f"    Added item {len(items)}: '{description[:50]}...' | NSN: {nsn} | Qty: {qty}")
    
    except Exception as e:
        print(f"ERROR: {e}")
        import traceback
        traceback.print_exc()
        return []
    
    print(f"\n=== Total items extracted: {len(items)} ===")
    return items


def generate_dd1750_from_pdf(bom_path, template_path, output_path, start_page=0):
    """Generate DD1750 - Simple and reliable."""
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
        import traceback
        traceback.print_exc()
        # Return empty template
        reader = PdfReader(template_path)
        writer = PdfWriter()
        writer.add_page(reader.pages[0])
        with open(output_path, 'wb') as f:
            writer.write(f)
        return output_path, 0
