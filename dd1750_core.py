"""DD1750 core - Items as image overlay."""

import io
import math
import re
from dataclasses import dataclass
from typing import List

import pdfplumber
from pypdf import PdfReader, PdfWriter
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import letter
from reportlab.lib.utils import ImageReader


@dataclass
class BomItem:
    line_no: int
    description: str
    nsn: str
    qty: int


def extract_items_from_pdf(pdf_path: str) -> List[BomItem]:
    items = []
    with pdfplumber.open(pdf_path) as pdf:
        for page in pdf.pages:
            tables = page.extract_tables()
            for table in tables:
                if len(table) < 2:
                    continue
                header = table[0]
                lv_idx = desc_idx = mat_idx = qty_idx = -1
                for i, cell in enumerate(header):
                    if cell:
                        text = str(cell).upper()
                        if 'LV' in text or 'LEVEL' in text:
                            lv_idx = i
                        elif 'DESC' in text:
                            desc_idx = i
                        elif 'MATERIAL' in text:
                            mat_idx = i
                        elif 'AUTH' in text and 'QTY' in text:
                            qty_idx = i
                if lv_idx == -1 or desc_idx == -1:
                    continue
                for row in table[1:]:
                    if not any(cell for cell in row if cell):
                        continue
                    lv_cell = row[lv_idx] if lv_idx < len(row) else None
                    if not lv_cell or str(lv_cell).strip().upper() != 'B':
                        continue
                    desc_cell = row[desc_idx] if desc_idx < len(row) else None
                    description = ""
                    if desc_cell:
                        lines = str(desc_cell).strip().split('\n')
                        description = lines[1].strip() if len(lines) >= 2 else lines[0].strip()
                        if '(' in description:
                            description = description.split('(')[0].strip()
                        description = re.sub(r'\s+', ' ', description).strip()
                    if not description:
                        continue
                    nsn = ""
                    if mat_idx > -1 and mat_idx < len(row):
                        mat_cell = row[mat_idx]
                        if mat_cell:
                            m = re.search(r'\b(\d{9})\b', str(mat_cell))
                            if m:
                                nsn = m.group(1)
                    qty = 1
                    if qty_idx > -1 and qty_idx < len(row):
                        qty_cell = row[qty_idx]
                        if qty_cell:
                            try:
                                qty = int(str(qty_cell).strip())
                            except:
                                pass
                    items.append(BomItem(len(items) + 1, description[:100], nsn, qty))
    return items


def generate_dd1750_from_pdf(bom_path: str, template_path: str, out_path: str):
    items = extract_items_from_pdf(bom_path)
    if not items:
        return out_path, 0
    
    # Read template
    template = PdfReader(template_path)
    template_page = template.pages[0]
    width = float(template_page.mediabox.width)
    height = float(template_page.mediabox.height)
    
    # Create items page as PDF
    items_packet = io.BytesIO()
    c = canvas.Canvas(items_packet, pagesize=(width, height))
    
    # Items table area only (below Y=616)
    first_row = 611.0
    
    for i, item in enumerate(items):
        y = first_row - (i * 27.25)
        if y < 90:  # Stop at bottom of table
            break
        
        c.setFont("Helvetica", 8)
        c.drawCentredString(66, y - 7, str(item.line_no))
        c.drawString(92, y - 7, item.description[:50])
        
        if item.nsn:
            c.setFont("Helvetica", 6)
            c.drawString(92, y - 12, f"NSN: {item.nsn}")
        
        c.setFont("Helvetica", 8)
        c.drawCentredString(386, y - 7, "EA")
        c.drawCentredString(431, y - 7, str(item.qty))
        c.drawCentredString(484, y - 7, "0")
        c.drawCentredString(540, y - 7, str(item.qty))
    
    c.save()
    items_packet.seek(0)
    
    # Save items as temp PDF and read as image
    with open('/tmp/items.pdf', 'wb') as f:
        f.write(items_packet.getvalue())
    
    # Create overlay as XObject
    overlay_reader = PdfReader('/tmp/items.pdf')
    overlay_page = overlay_reader.pages[0]
    
    # Get template content
    content = overlay_page.get_contents()
    if content:
        template_page.merge_page(overlay_page)
    
    writer = PdfWriter()
    writer.add_page(template_page)
    
    with open(out_path, 'wb') as f:
        writer.write(f)
    
    return out_path, len(items)
