import os
import tempfile
from flask import Flask, render_template, request, send_file, flash, redirect, url_for
from werkzeug.utils import secure_filename
from dd1750_core import generate_dd1750_from_pdf

app = Flask(__name__)
app.secret_key = os.environ.get('SECRET_KEY', 'dev-key-change-in-production')
app.config['MAX_CONTENT_LENGTH'] = 200 * 1024 * 1024  # 200MB

ALLOWED_EXTENSIONS = {'pdf'}

def allowed_file(filename):
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/generate', methods=['POST'])
def generate():
    if 'bom_file' not in request.files or 'template_file' not in request.files:
        flash('Both BOM PDF and template PDF are required')
        return redirect(url_for('index'))
    
    bom_file = request.files['bom_file']
    template_file = request.files['template_file']
    
    if bom_file.filename == '' or template_file.filename == '':
        flash('Both files must be selected')
        return redirect(url_for('index'))
    
    if not (allowed_file(bom_file.filename) and allowed_file(template_file.filename)):
        flash('Both files must be PDF format')
        return redirect(url_for('index'))
    
    try:
        start_page = int(request.form.get('start_page', 0))
    except:
        start_page = 0
    
    try:
        with tempfile.TemporaryDirectory() as temp_dir:
            bom_path = os.path.join(temp_dir, secure_filename(bom_file.filename))
            template_path = os.path.join(temp_dir, secure_filename(template_file.filename))
            output_path = os.path.join(temp_dir, 'DD1750_filled.pdf')
            
            bom_file.save(bom_path)
            template_file.save(template_path)
            
            output_path, item_count = generate_dd1750_from_pdf(
                bom_pdf_path=bom_path,
                template_pdf_path=template_path,
                out_pdf_path=output_path,
                start_page=start_page
            )
            
            if item_count == 0:
                flash('No items found in BOM PDF. Please check your file.')
                return redirect(url_for('index'))
            
            filename = f"DD1750_{item_count}_items.pdf"
            return send_file(output_path, as_attachment=True, download_name=filename)
            
    except Exception as e:
        flash(f'Error processing files: {str(e)}')
        return redirect(url_for('index'))

if __name__ == '__main__':
    port = int(os.environ.get('PORT', 8000))
    app.run(host='0.0.0.0', port=port, debug=False)
