"""DD1750 core - Only extract rows where LV = "B"."""

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
    """Extract items where LV = 'B' only."""
    items = []
    
    try:
        with pdfplumber.open(pdf_path) as pdf:
            for page in pdf.pages[start_page:]:
                # Try to extract tables first (more reliable)
                tables = page.extract_tables()
                
                if tables:
                    for table in tables:
                        if len(table) > 1:  # Has at least header + data
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
                            
                            # Process rows
                            for row in table[1:]:
                                # Check if this row has LV = 'B'
                                if 'lv' in col_indices and len(row) > col_indices['lv']:
                                    lv_value = str(row[col_indices['lv']]).strip().upper()
                                    
                                    if lv_value == 'B':
                                        # Extract description
                                        description = ""
                                        if 'desc' in col_indices and len(row) > col_indices['desc']:
                                            desc_cell = row[col_indices['desc']]
                                            if desc_cell:
                                                description = str(desc_cell).strip()
                                                # Clean description
                                                if '(' in description:
                                                    description = description.split('(')[0].strip()
                                                if ':' in description:
                                                    description = description.split(':')[0].strip()
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
                                        
                                        if description:  # Only add if we have a description
                                            items.append(BomItem(
                                                line_no=len(items) + 1,
                                                description=description[:100],
                                                nsn=nsn,
                                                qty=qty
                                            ))
                
                # Fallback: If no tables found, try text extraction
                if not items:
                    text = page.extract_text() or ""
                    lines = [ln.strip() for ln in text.split('\n') if ln.strip()]
                    
                    i = 0
                    while i < len(lines):
                        line = lines[i]
                        
                        # Look for LV = B pattern
                        if re.match(r'^B\s+', line) or (' B ' in line and len(line.split()) >= 3):
                            # Extract description (skip first 2 columns: Material and LV)
                            parts = line.split()
                            if len(parts) >= 3:
                                # Join remaining parts as description
                                desc_parts = parts[2:]
                                description = ' '.join(desc_parts)
                                
                                # Clean description
                                if '(' in description:
                                    description = description.split('(')[0].strip()
                                if ':' in description:
                                    description = description.split(':')[0].strip()
                                description = re.sub(r'\s+', ' ', description).strip()
                                
                                # Look for NSN in current or nearby lines
                                nsn = ""
                                for j in range(max(0, i-1), min(len(lines), i+2)):
                                    nsn_match = re.search(r'\b(\d{9})\b', lines[j])
                                    if nsn_match:
                                        nsn = nsn_match.group(1)
                                        break
                                
                                # Look for quantity
                                qty = 1
                                qty_match = re.search(r'\b(\d+)\s*$', line)
                                if qty_match:
                                    try:
                                        qty = int(qty_match.group(1))
                                    except:
                                        qty = 1
                                
                                if description:
                                    items.append(BomItem(
                                        line_no=len(items) + 1,
                                        description=description[:100],
                                        nsn=nsn,
                                        qty=qty
                                    ))
                        
                        i += 1
    
    except Exception as e:
        print(f"ERROR in extraction: {e}")
        return []
    
    return items


def generate_dd1750_from_pdf(bom_path, template_path, output_path, start_page=0):
    """Generate DD1750 - Only LV = 'B' items."""
    try:
        items = extract_items_from_pdf(bom_path, start_page)
        
        print(f"DEBUG: Found {len(items)} items with LV = 'B'")
        for i, item in enumerate(items[:5], 1):
            print(f"  Item {i}: '{item.description}' | NSN: {item.nsn} | Qty: {item.qty}")
        
        if not items:
            # Return empty template
            reader = PdfReader(template_path)
            writer = PdfWriter()
            writer.add_page(reader.pages[0])
            with open(output_path, 'wb') as f:
                writer.write(f)
            return output_path, 0
        
        # Create overlay
        packet = io.BytesIO()
        can = canvas.Canvas(packet, pagesize=(PAGE_W, PAGE_H))
        
        first_row_top = Y_TABLE_TOP_LINE - 5.0
        max_w = (X_CONTENT_R - X_CONTENT_L) - 2 * PAD_X
        
        for i in range(min(len(items), ROWS_PER_PAGE)):
            item = items[i]
            y = first_row_top - (i * ROW_H)
            y_desc = y - 7.0
            y_nsn = y - 12.2
            
            # Box number
            can.setFont("Helvetica", 8)
            can.drawCentredString((X_BOX_L + X_BOX_R)/2, y_desc, str(item.line_no))
            
            # Description
            can.setFont("Helvetica", 7)
            desc = item.description[:50] if len(item.description) > 50 else item.description
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
        reader = PdfReader(template_path)
        writer = PdfWriter()
        
        overlay = PdfReader(packet)
        page = reader.pages[0]
        page.merge_page(overlay.pages[0])
        writer.add_page(page)
        
        with open(output_path, 'wb') as f:
            writer.write(f)
        
        return output_path, len(items)
        
    except Exception as e:
        print(f"CRITICAL ERROR: {e}")
        # Return empty template on any error
        reader = PdfReader(template_path)
        writer = PdfWriter()
        writer.add_page(reader.pages[0])
        with open(output_path, 'wb') as f:
            writer.write(f)
        return output_path, 0
