"""DD1750 core - Direct PDF Text Injection for Admin Fields."""

import io
import math
import re
from dataclasses import dataclass
from typing import List, Dict

import pdfplumber
from pypdf import PdfReader, PdfWriter, Transformation, ContentType
from reportlab.pdfgen import canvas
from reportlab.pdfbase import pdfmetrics
from reportlab.lib.pagesizes import letter
from reportlab.lib.utils import ImageReader


ROWS_PER_PAGE = 18
PAGE_W, PAGE_H = letter

# Column positions
X_BOX_L, X_BOX_R = 44.0, 88.0
X_CONTENT_L, X_CONTENT_R = 88.0, 365.0
X_UOI_L, X_UOI_R = 365.0, 408.5
X_INIT_L, X_INIT_R = 408.5, 453.5
X_SPARES_L, X_SPARES_R = 453.5, 514.5
X_TOTAL_L, X_TOTAL_R = 514.5, 566.0

Y_TABLE_TOP = 616.0
Y_TABLE_BOTTOM = 89.5
ROW_H = (Y_TABLE_TOP - Y_TABLE_BOTTOM) / ROWS_PER_PAGE
PAD_X = 3.0


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
                    lv_idx = desc_idx = mat_idx = auth_idx = -1
                    
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
                                auth_idx = i
                    
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
                            description = re.sub(r'\s+(WTY|ARC|CIIC|UI|SCMC|EA|AY|9K|9G)$', '', description, flags=re.IGNORECASE)
                            description = re.sub(r'\s+', ' ', description).strip()
                        
                        if not description:
                            continue
                        
                        nsn = ""
                        if mat_idx > -1 and mat_idx < len(row):
                            mat_cell = row[mat_idx]
                            if mat_cell:
                                match = re.search(r'\b(\d{9})\b', str(mat_cell))
                                if match:
                                    nsn = match.group(1)
                        
                        qty = 1
                        if auth_idx > -1 and auth_idx < len(row):
                            qty_cell = row[auth_idx]
                            if qty_cell:
                                match = re.search(r'(\d+)', str(qty_cell))
                                if match:
                                    qty = int(match.group(1))
                        
                        items.append(BomItem(len(items) + 1, description[:100], nsn, qty))
    
    except Exception as e:
        print(f"ERROR: {e}")
        return []
    
    return items


def fill_admin_fields(template_pdf_path: str, admin_data: Dict) -> str:
    """Fill admin fields directly into the template PDF using pypdf update_page.
    This preserves form fields and adds the filled text."""
    try:
        reader = PdfReader(template_pdf_path)
        writer = PdfReader()
        
        # Copy all pages to writer
        for i, page in enumerate(reader.pages):
            writer.add_page(page)
        
        # Update form fields directly on ALL pages (this is key for multi-page PDFs)
        # We use set_fields to populate the form fields with admin data
        
        field_mappings = {
            'unit': admin_data.get('unit', ''),
            'date': admin_data.get('date', ''),
            'requisition': admin_data.get('requisition_no', ''),
            'order': admin_data.get('order_no', ''),
            'boxes': admin_data.get('num_boxes', ''),
            'packed_by': admin_data.get('packed_by', ''),
            'end_item': admin_data.get('end_item', ''),
            'model': admin_data.get('model', ''),
        }
        
        # Get all form fields from the first page
        # We check all pages to be safe (in case template has different fields)
        all_fields = {}
        for i, page in enumerate(writer.pages):
            try:
                if '/Annots' in page:
                    annots = page['/Annots']
                    for annot in annots:
                        if annot.get_object() == '/T' or annot.get_object() == '/Tx':
                            field_name = annot.get('/T') or annot.get_object().get('/TU')
                            # Use common field name mappings
                            if 'unit' in field_name.lower():
                                all_fields['unit'] = field_name
                            elif 'date' in field_name.lower():
                                all_fields['date'] = field_name
                            elif 'requisition' in field_name.lower():
                                all_fields['requisition'] = field_name
                            elif 'order' in field_name.lower():
                                all_fields['order'] = field_name
                            elif 'box' in field_name.lower():
                                all_fields['boxes'] = field_name
                            elif 'packed' in field_name.lower():
                                all_fields['packed_by'] = field_name
                            elif 'end' in field_name.lower() or 'item' in field_name.lower():
                                all_fields['end_item'] = field_name
                            elif 'model' in field_name.lower():
                                all_fields['model'] = field_name
            except:
                pass
        
        # Set the field values for all pages
        for i, page in enumerate(writer.pages):
            try:
                if '/Annots' in page:
                    annots = page['/Annots']
                    
                    for annot in annots:
                        obj = annot.get_object()
                        if obj and obj.get('/FT') == '/Ft':  # Text field
                            field_name = all_fields.get(obj.get('/T'))
                            if field_name:
                                annot.update({
                                    '/T': obj.get('/T'),
                                    '/FT': '/Ft' + admin_data.get(field_name, ''),
                                    '/V': admin_data.get(field_name, '')
                                })
                        elif obj and obj.get_object().get('/FT') == '/Tx':  # Checkbox/Radio
                            # Set as needed, but for now skip
                            pass
            except:
                pass
        
        # Write to a new PDF
        output_path = template_pdf_path.replace('.pdf', '_filled.pdf')
        writer.write(output_path)
        
        return output_path
    
    except Exception as e:
        print(f"ERROR filling admin: {e}")
        return template_pdf_path


def generate_dd1750_from_pdf(bom_path: str, template_path: str, out_path: str, start_page: int = 0, admin_data: Dict = None):
    """Generate DD1750 with admin fields filled in template."""
    
    if admin_data is None:
        admin_data = {}
    
    try:
        # Step 1: Fill admin fields in template (creates a new PDF with filled data)
        filled_template_path = fill_admin_fields(template_path, admin_data)
        
        # Step 2: Generate items from BOM
        items = extract_items_from_pdf(bom_path, start_page)
        
        print(f"Items found: {len(items)}")
        
        if not items:
            return out_path, 0
        
        # Step 3: Use the FILLED template as the base
        template_reader = PdfReader(filled_template_path)
        
        total_pages = math.ceil(len(items) / ROWS_PER_PAGE)
        writer = PdfWriter()
        
        for page_num in range(total_pages):
            start_idx = page_num * ROWS_PER_PAGE
            end_idx = min((page_num + 1) * ROWS_PER_PAGE, len(items))
            page_items = items[start_idx:end_idx]
            
            # Create items overlay
            packet = io.BytesIO()
            c = canvas.Canvas(packet, pagesize=letter)
            
            first_row = Y_TABLE_TOP - 5.0
            
            for i, item in enumerate(page_items):
                y = first_row - (i * ROW_H)
                
                c.setFont("Helvetica", 8)
                c.drawCentredString((X_BOX_L + X_BOX_R) / 2, y - 7, str(item.line_no))
                
                c.setFont("Helvetica", 7)
                c.drawString(X_CONTENT_L + PAD_X, y - 7, item.description[:50])
                
                if item.nsn:
                    c.setFont("Helvetica", 6)
                    c.drawString(X_CONTENT_L + PAD_X, y - 12, f"NSN: {item.nsn}")
                
                c.setFont("Helvetica", 8)
                c.drawCentredString((X_UOI_L + X_UOI_R) / 2, y - 7, "EA")
                c.drawCentredString((X_INIT_L + X_INIT_R) / 2, y - 7, str(item.qty))
                c.drawCentredString((X_SPARES_L + X_SPARES_R) / 2, y - 7, "0")
                c.drawCentredString((X_TOTAL_L + X_TOTAL_R) / 2, y - 7, str(item.qty))
            
            c.save()
            packet.seek(0)
            
            overlay = PdfReader(packet)
            
            # MERGE ITEMS INTO FILLED TEMPLATE PAGE
            # The filled template already has admin data filled in.
            # We are just adding the items overlay on top of it.
            # Since we are using update_page on the same PdfWriter object,
            # the form fields should be preserved!
            page = template_reader.pages[page_num]
            page.merge_page(overlay.pages[0])
            writer.add_page(page)
        
        with open(out_path, 'wb') as f:
            writer.write(f)
        
        return out_path, len(items)
        
    except Exception as e:
        print(f"CRITICAL ERROR: {e}")
        import traceback
        traceback.print_exc()
        
        # Fallback to blank template if error
        try:
            reader = PdfReader(template_path)
            writer = PdfWriter()
            writer.add_page(reader.pages[0])
            with open(out_path, 'wb') as f:
                writer.write(f)
        except:
            pass
        
        return out_path, 0
