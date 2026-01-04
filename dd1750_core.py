def generate_dd1750_from_pdf(bom_path, template_path, out_path, start_page=0, admin_data=None):
    items = extract_items_from_pdf(bom_path)
    
    if not items:
        return out_path, 0
    
    total_pages = math.ceil(len(items) / ROWS_PER_PAGE)
    writer = PdfWriter()
    
    # Read template ONCE (outside the loop)
    template_reader = PdfReader(template_path)
    first_page = template_reader.pages[0]
    
    for page_num in range(total_pages):
        start_idx = page_num * ROWS_PER_PAGE
        end_idx = min((page_num + 1) * ROWS_PER_PAGE, len(items))
        page_items = items[start_idx:end_idx]
        
        # Create overlay
        packet = io.BytesIO()
        c = canvas.Canvas(packet, pagesize=letter)
        first_row = Y_TABLE_TOP - 5.0
        
        for i, item in enumerate(page_items):
            y = first_row - (i * ROW_H)
            
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
        
        # Draw admin fields ONLY on first page of output PDF
        if page_num == 0 and admin_data:
            # Draw admin fields at the top
            if admin_data.get('unit'):
                c.setFont("Helvetica", 10)
                c.drawString(50, 745, admin_data['unit'][:30])
            
            if admin_data.get('requisition_no'):
                c.setFont("Helvetica", 10)
                c.drawString(250, 745, admin_data['requisition_no'][:30])
            
            if admin_data.get('date'):
                c.setFont("Helvetica", 10)
                c.drawString(50, 715, admin_data['date'])
            
            if admin_data.get('order_no'):
                c.setFont("Helvetica", 10)
                c.drawString(250, 715, admin_data['order_no'][:30])
            
            if admin_data.get('num_boxes'):
                c.setFont("Helvetica", 10)
                c.drawString(450, 715, admin_data['num_boxes'][:5])
            
            # Packed By (bottom section)
            if admin_data.get('packed_by'):
                c.setFont("Helvetica", 10)
                c.drawString(44, 130, f"PACKED BY: {admin_data['packed_by']}")
        
        c.save()
        packet.seek(0)
        
        overlay = PdfReader(packet)
        
        # Use the FIRST page of the template for the first output page
        # Use the SAME page for subsequent output pages
        if page_num == 0:
            page = first_page
        else:
            # For pages 2+, use the FIRST page of the template
            # (This preserves the empty admin fields on all pages)
            page = template_reader.pages[0]
        
        page.merge_page(overlay.pages[0])
        writer.add_page(page)
    
    with open(output_path, 'wb') as f:
        writer.write(f)
    
    return output_path, len(items)
