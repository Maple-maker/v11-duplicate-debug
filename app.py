import os
import tempfile
from datetime import datetime
from flask import Flask, render_template, request, send_file, flash, redirect
from werkzeug.utils import secure_filename
from dd1750_core import generate_dd1750_from_pdf

app = Flask(__name__)
app.secret_key = 'dev-key-change-in-production'
app.config['MAX_CONTENT_LENGTH'] = 200 * 1024 * 1024

ALLOWED_EXTENSIONS = {'pdf'}

def allowed_file(filename):
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS

@app.route('/')
def index():
    current_date = datetime.now().strftime('%Y-%m-%d')
    return render_template('index.html', current_date=current_date)

@app.route('/generate', methods=['POST'])
def generate():
    if 'bom_file' not in request.files or 'template_file' not in request.files:
        flash('Both files required')
        return redirect('/')
    
    bom_file = request.files['bom_file']
    template_file = request.files['template_file']
    
    if not allowed_file(bom_file.filename) or not allowed_file(template_file.filename):
        flash('Both must be PDFs')
        return redirect('/')
    
    try:
        start_page = int(request.form.get('start_page', 0))
        
        admin_data = {
            'unit': request.form.get('unit', '').strip(),
            'packed_by': request.form.get('packed_by', '').strip(),
            'num_boxes': request.form.get('num_boxes', '').strip(),
            'requisition_no': request.form.get('requisition_no', '').strip(),
            'order_no': request.form.get('order_no', '').strip(),
            'date': request.form.get('date', datetime.now().strftime('%Y-%m-%d')).strip(),
        }
        
        print(f"DEBUG: Admin data received: {admin_data}")
        
        with tempfile.TemporaryDirectory() as temp_dir:
            bom_path = os.path.join(temp_dir, 'bom.pdf')
            template_path = os.path.join(temp_dir, 'template.pdf')
            output_path = os.path.join(temp_dir, 'output.pdf')
            
            bom_file.save(bom_path)
            template_file.save(template_path)
            
            output_path, item_count = generate_dd1750_from_pdf(
                bom_path, template_path, output_path, start_page, admin_data
            )
            
            if item_count == 0:
                flash('No items found in BOM')
                return redirect('/')
            
            filename = f"DD1750_{admin_data['requisition_no'] or 'filled'}_{item_count}_items.pdf"
            return send_file(output_path, as_attachment=True, download_name=filename)
            
    except Exception as e:
        print(f"ERROR: {e}")
        flash(f'Error: {str(e)}')
        return redirect('/')

if __name__ == '__main__':
    port = int(os.environ.get('PORT', 8000))
    app.run(host='0.0.0.0', port=port)
