import os
import tempfile
from flask import Flask, render_template, request, send_file
from dd1750_core import generate_dd1750_from_pdf

app = Flask(__name__)
app.secret_key = os.environ.get('SECRET_KEY', 'dd1750-secret-key')
app.config['MAX_CONTENT_LENGTH'] = 200 * 1024 * 1024  # 200MB

ALLOWED_EXTENSIONS = {'pdf'}

def allowed_file(filename):
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/generate', methods=['POST'])
def generate():
    if 'bom_file' not in request.files:
        return render_template('index.html')
    
    if 'template_file' not in request.files:
        return render_template('index.html')
    
    bom_file = request.files['bom_file']
    template_file = request.files['template_file']
    
    if bom_file.filename == '' or template_file.filename == '':
        return render_template('index.html')
    
    if not (allowed_file(bom_file.filename) and allowed_file(template_file.filename)):
        return render_template('index.html')
    
    try:
        start_page = int(request.form.get('start_page', 0))
        
        with tempfile.TemporaryDirectory() as tmpdir:
            bom_path = os.path.join(tmpdir, 'bom.pdf')
            template_path = os.path.join(tmpdir, 'template.pdf')
            output_path = os.path.join(tmpdir, 'DD1750.pdf')
            
            bom_file.save(bom_path)
            template_file.save(template_path)
            
            # Pass start_page
            out_path, count = generate_dd1750_from_pdf(
                bom_pdf_path=bom_path,
                template_pdf_path=template_path,
                out_pdf_path=output_path,
                start_page=start_page
            )
            
            # Check file exists before sending
            if not os.path.exists(out_path):
                print(f"ERROR: Output file not created at {output_path}")
                return render_template('index.html', error="Processing failed. Please check logs.")
            
            if count == 0:
                print("DEBUG: count is 0, but file exists. Returning anyway.")
                return send_file(out_path, as_attachment=True, download_name='DD1750.pdf')
            
            return send_file(out_path, as_attachment=True, download_name='DD1750.pdf')
    
    except Exception as e:
        print(f"ERROR: {e}")
        import traceback
        traceback.print_exc()
        return render_template('index.html', error=f"Error: {str(e)}")

if __name__ == '__main__':
    port = int(os.environ.get('PORT', 8000))
    app.run(host='0.0.0.0', port=port)
