import os
import tempfile
from flask import Flask, render_template, request, send_file
from dd1750_core import generate_dd1750_from_pdf

app = Flask(__name__)

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
    
    if not (bom_file.filename.lower().endswith('.pdf') and template_file.filename.lower().endswith('.pdf')):
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
            
            print(f"DEBUG: Saving BOM: {bom_file.filename}")
            print(f"DEBUG: Saving Template: {template_file.filename}")
            print(f"DEBUG: Output path: {out_path}")
            
            bom_file.save(bom_path)
            template_file.save(tpl_path)
            
            # Generate
            out_path, count = generate_dd1750_from_pdf(
                bom_path=bom_path,
                template_path=tpl_path,
                out_path=out_path,
                start_page=start_page
            )
            
            # Dumb Mode: Send file even if count is 0
            return send_file(out_path, as_attachment=True, download_name='DD1750.pdf')
    
    except Exception as e:
        print(f"CRITICAL ERROR: {e}")
        import traceback
        traceback.print_exc()
        return render_template('index.html', error=f"Error: {str(e)}")

if __name__ == '__main__':
    port = int(os.environ.get('PORT', 8000))
    app.run(host='0.0.0.0', port=port)
