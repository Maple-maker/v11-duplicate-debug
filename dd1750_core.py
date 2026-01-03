"""DD1750 core - Minimal crash-proof version."""

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
    """Extract items - SIMPLEST possible version."""
    items = []
    
    try:
        with pdfplumber.open(pdf_path) as pdf:
            for page in pdf.pages[start_page:]:
                text = page.extract_text() or ""
                lines = [ln.strip() for ln in text.split('\n') if ln.strip()]
                
                i = 0
                while i < len(lines):
                    line = lines[i]
                    
                    # Skip headers
                    if line.startswith("COEI-") or line.startswith("BII-"):
                        i += 1
                        continue
                    
                    # Look for item descriptions (contains comma)
                    if ',' in line and len(line) > 5:
                        # Extract description up to colon or parenthesis
                        desc = line
                        if ':' in desc:
                            desc = desc.split(':')[0]
                        if '(' in desc:
                            desc = desc.split('(')[0]
                        
                        desc = desc.strip()
                        
                        # Remove unwanted codes
                        desc = re.sub(r'\b(WTY|ARC|CIIC|UI|SCMC|EA|AY|9K|9G)\b', '', desc, flags=re.IGNORECASE)
                        desc = re.sub(r'\s+', ' ', desc).strip()
                        
                        if desc and len(desc) > 3:
                            # Find NSN
                            nsn = ""
                            for j in range(max(0, i-2), min(len(lines), i+3)):
                                nsn_match = re.search(r'\b(\d{9})\b', lines[j])
                                if nsn_match:
                                    nsn = nsn_match.group(1)
                                    break
                            
                            # Find quantity
                            qty = 1
                            qty_match = re.search(r'\b(\d+)\s*$', line)
                            if qty_match:
                                try:
                                    qty = int(qty_match.group(1))
                                except:
                                    qty = 1
                            
                            items.append(BomItem(
                                line_no=len(items) + 1,
                                description=desc[:100],  # Limit length
                                nsn=nsn,
                                qty=qty
                            ))
                    
                    i += 1
    except:
        # Return empty list on any error
        return []
    
    return items


def generate_dd1750_from_pdf(bom_path, template_path, output_path, start_page=0):
    """Generate DD1750 - crash-proof."""
    try:
        items = extract_items_from_pdf(bom_path, start_page)
        
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
            can.drawString(X_CONTENT_L + PAD_X, y_desc, item.description[:50])
            
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
        # Return empty template on any error
        reader = PdfReader(template_path)
        writer = PdfWriter()
        writer.add_page(reader.pages[0])
        with open(output_path, 'wb') as f:
            writer.write(f)
        return output_path, 0
