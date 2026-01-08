import os
import sys
import tempfile
from flask import Flask, render_template, request, send_file
from dd1750_core import generate_dd1750_from_pdf

app = Flask(__name__)
app.secret_key = os.environ.get('SECRET_KEY', 'dd1750-secret-key')
app.config['MAX_CONTENT_LENGTH'] = 200 * 1024 * 1024

ALLOWED_EXTENSIONS = {'pdf'}

def allowed_file(filename):
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/generate', methods=['POST'])
def generate():
    if 'bom_file' not in request.files:
        return render_template('index.html', error='No BOM PDF uploaded')
    
    if 'template_file' not in request.files:
        return render_template('index.html', error='No DD1750 Template PDF uploaded')
    
    bom_file = request.files['bom_file']
    template_file = request.files['template_file']
    
    if bom_file.filename == '' or template_file.filename == '':
        return render_template('index.html', error='Both files must be selected')
    
    if not (allowed_file(bom_file.filename) and allowed_file(template_file.filename)):
        return render_template('index.html', error='Both files must be PDF format')
    
    try:
        start_page = int(request.form.get('start_page', 0))
    except:
        start_page = 0
    
    try:
        with tempfile.TemporaryDirectory() as tmpdir:
            bom_path = os.path.join(tmpdir, 'bom.pdf')
            tpl_path = os.path.join(tmpdir, 'template.pdf')
            out_path = os.path.join(tmpdir, 'DD1750.pdf')
            
            bom_file.save(bom_path)
            template_file.save(tpl_path)
            
            print(f"DEBUG: Processing BOM: {bom_file.filename}, Template: {template_file.filename}, Output: {out_path}")
            
            out_path, count = generate_dd1750_from_pdf(
                bom_pdf_path=bom_path,
                template_pdf_path=tpl_path,
                out_pdf_path=out_path,
                start_page=start_page
            )
            
            print(f"DEBUG: Generation complete. Count: {count}, Output path: {out_path}")
            
            if not os.path.exists(out_path):
                print("ERROR: Output file does not exist!")
                return render_template('index.html', error='Internal error: PDF could not be generated')
            
            file_size = os.path.getsize(out_path)
            print(f"DEBUG: Output file size: {file_size} bytes")
            
            if file_size == 0:
                print("ERROR: Output file is 0 bytes!")
                return render_template('index.html', error='Internal error: Generated PDF is empty')
            
            print(f"DEBUG: Sending file to user...")
            
            return send_file(out_path, as_attachment=True, download_name='DD1750.pdf')
    
    except Exception as e:
        print(f"CRITICAL ERROR: {e}", file=sys.stderr)
        import traceback
        traceback.print_exc()
        return render_template('index.html', error=f"Error: {str(e)}")

if __name__ == '__main__':
    port = int(os.environ.get('PORT', 8000))
    app.run(host='0.0.0.0', port=port)
