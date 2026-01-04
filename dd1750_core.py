"""DD1750 core - Simple and reliable extraction."""

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
    """Simple, reliable extraction focusing on core requirements."""
    items = []
    
    try:
        with pdfplumber.open(pdf_path) as pdf:
            for page in pdf.pages[start_page:]:
                # Get all text
                text = page.extract_text() or ""
                if not text:
                    continue
                
                # Split into lines
                lines = [ln.strip() for ln in text.split('\n') if ln.strip()]
                
                i = 0
                while i < len(lines):
                    line = lines[i]
                    
                    # DEBUG: Print line to see what we're processing
                    # print(f"Line {i}: {line}")
                    
                    # Pattern 1: Line starts with material code, then B, then description
                    # Example: "011800996 B BASE ASSEMBLY, OUTRIGGER"
                    if re.match(r'^\d{9}\s+B\s+', line):
                        parts = line.split()
                        if len(parts) >= 3:
                            # NSN is first part
                            nsn = parts[0]
                            
                            # Description starts after "B"
                            b_index = parts.index('B') if 'B' in parts else 1
                            desc_parts = parts[b_index + 1:]
                            
                            # Join description parts
                            description = ' '.join(desc_parts)
                            
                            # Extract only top text (stop at colon or parenthesis)
                            if ':' in description:
                                description = description.split(':')[0].strip()
                            if '(' in description:
                                description = description.split('(')[0].strip()
                            
                            # Remove any WTY/ARC/CIIC/UI/SCMC codes
                            description = re.sub(r'\b(WTY|ARC|CIIC|UI|SCMC|EA|AY|9K|9G)\b', '', description, flags=re.IGNORECASE)
                            description = re.sub(r'\s+', ' ', description).strip()
                            
                            # Extract quantity (usually at end)
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
                    
                    # Pattern 2: Line has "B" somewhere and looks like an item
                    # Example: "B BASE ASSEMBLY, OUTRIGGER" or contains comma
                    elif ' B ' in line and ',' in line:
                        parts = line.split()
                        
                        # Find the "B"
                        if 'B' in parts:
                            b_index = parts.index('B')
                            
                            # Check if part before B could be NSN
                            nsn = ""
                            if b_index > 0 and re.match(r'^\d{9}$', parts[b_index - 1]):
                                nsn = parts[b_index - 1]
                            
                            # Description starts after B
                            desc_parts = parts[b_index + 1:]
                            description = ' '.join(desc_parts)
                            
                            # Extract only top text
                            if ':' in description:
                                description = description.split(':')[0].strip()
                            if '(' in description:
                                description = description.split('(')[0].strip()
                            
                            # Clean description
                            description = re.sub(r'\b(WTY|ARC|CIIC|UI|SCMC|EA|AY|9K|9G)\b', '', description, flags=re.IGNORECASE)
                            description = re.sub(r'\s+', ' ', description).strip()
                            
                            # Extract quantity
                            qty = 1
                            if description and description.split()[-1].isdigit():
                                qty = int(description.split()[-1])
                                description = ' '.join(description.split()[:-1])
                            
                            # Look for NSN in nearby lines if not found
                            if not nsn:
                                for j in range(max(0, i-2), min(len(lines), i+3)):
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
                    
                    # Pattern 3: Line that looks like a description (contains comma, not a code)
                    elif ',' in line and len(line) > 10 and not line.startswith('COEI-') and not line.startswith('BII-'):
                        # Check if previous line has NSN
                        nsn = ""
                        if i > 0 and re.match(r'^\d{9}$', lines[i-1]):
                            nsn = lines[i-1]
                        
                        # Check if line starts with B (might be missing space)
                        description = line
                        if description.startswith('B '):
                            description = description[2:].strip()
                        
                        # Extract only top text
                        if ':' in description:
                            description = description.split(':')[0].strip()
                        if '(' in description:
                            description = description.split('(')[0].strip()
                        
                        # Clean description
                        description = re.sub(r'\b(WTY|ARC|CIIC|UI|SCMC|EA|AY|9K|9G)\b', '', description, flags=re.IGNORECASE)
                        description = re.sub(r'\s+', ' ', description).strip()
                        
                        # Extract quantity
                        qty = 1
                        if description and description.split()[-1].isdigit():
                            qty = int(description.split()[-1])
                            description = ' '.join(description.split()[:-1])
                        
                        # Look for NSN in nearby lines if not found
                        if not nsn:
                            for j in range(max(0, i-3), min(len(lines), i+3)):
                                nsn_match = re.search(r'\b(\d{9})\b', lines[j])
                                if nsn_match:
                                    nsn = nsn_match.group(1)
                                    break
                        
                        if description and len(description) > 3:
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
    """Generate DD1750 - Simple and reliable."""
    try:
        items = extract_items_from_pdf(bom_path, start_page)
        
        print(f"=== DEBUG: Found {len(items)} items ===")
        for i, item in enumerate(items, 1):
            print(f"Item {i}: '{item.description}' | NSN: {item.nsn} | Qty: {item.qty}")
        print("=== END DEBUG ===")
        
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
