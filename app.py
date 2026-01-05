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
        return render_template('index.html', error="No BOM PDF uploaded")
    
    if 'template_file' not in request.files:
        return render_template('index.html', error="No DD1750 Template PDF uploaded")
    
    bom_file = request.files['bom_file']
    template_file = request.files['template_file']
    
    if bom_file.filename == '' or template_file.filename == '':
        return render_template('index.html', error="Both files must be selected")
    
    try:
        with tempfile.TemporaryDirectory() as tmpdir:
            bom_path = os.path.join(tmpdir, 'bom.pdf')
            tpl_path = os.path.join(tmpdir, 'template.pdf')
            out_path = os.path.join(tmpdir, 'DD1750.pdf')
            
            print(f"DEBUG: Saving BOM to {bom_path}")
            print(f"DEBUG: Saving template to {tpl_path}")
            print(f"DEBUG: Output path: {out_path}")
            
            bom_file.save(bom_path)
            template_file.save(tpl_path)
            
            print(f"DEBUG: Generating DD1750...")
            out_path, count = generate_dd1750_from_pdf(
                bom_pdf_path=bom_path,
                template_pdf_path=tpl_path,
                out_pdf_path=out_path
            )
            
            print(f"DEBUG: Generation complete. Count: {count}")
            print(f"DEBUG: Checking if output file exists...")
            
            if not os.path.exists(out_path):
                print(f"ERROR: Output file not found at {out_path}")
                return render_template('index.html', error="Internal error: PDF could not be generated")
            
            file_size = os.path.getsize(out_path)
            print(f"DEBUG: Output file size: {file_size} bytes")
            
            if file_size == 0:
                print("ERROR: Output file is empty (0 bytes)")
                return render_template('index.html', error="Internal error: Generated PDF is empty")
            
            print(f"DEBUG: Sending file to user...")
            return send_file(out_path, as_attachment=True, download_name='DD1750.pdf')
    
    except Exception as e:
        print(f"CRITICAL ERROR: {e}")
        import traceback
        traceback.print_exc()
        return render_template('index.html', error=f"Error: {str(e)}")

if __name__ == '__main__':
    port = int(os.environ.get('PORT', 8000))
    app.run(host='0.0.0.0', port=port)
