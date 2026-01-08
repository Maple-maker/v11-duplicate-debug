import os
import tempfile
import shutil
from flask import Flask, render_template, request, send_file
from dd1750_core import generate_dd1750_from_pdf

app = Flask(__name__)
app.secret_key = 'dd1750-secret-key-2026'

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
    
    try:
        # Create a persistent temporary directory (NOT using tempfile.TemporaryDirectory)
        # to ensure file stays on disk during send_file
        import uuid
        tmpdir = f"/tmp/dd1750-{uuid.uuid4()}"
        os.makedirs(tmpdir, exist_ok=True)
        
        bom_path = os.path.join(tmpdir, 'bom.pdf')
        tpl_path = os.path.join(tmpdir, 'template.pdf')
        
        # NOTE: Don't use out_path here. The core function will use a generated path.
        out_dir = tmpdir
        
        bom_file.save(bom_path)
        template_file.save(tpl_path)
        
        print(f"DEBUG: Processing BOM: {bom_path}")
        print(f"DEBUG: Processing Template: {tpl_path}")
        
        # Generate
        from pathlib import Path
        out_path = os.path.join(tmpdir, 'DD1750.pdf')
        
        count = generate_dd1750_from_pdf(
            bom_path=bom_path,
            template_path=tpl_path,
            out_path=out_path
        )
        
        print(f"DEBUG: Generated {count} items to {out_path}")
        
        if count == 0:
            return render_template('index.html', error='No items found in BOM')
        
        if not os.path.exists(out_path):
            print(f"ERROR: File does not exist at {out_path}")
            return render_template('index.html', error='Internal error: PDF could not be generated')
        
        file_size = os.path.getsize(out_path)
        print(f"DEBUG: File exists, Size: {file_size}")
        
        if file_size == 0:
            print("ERROR: File is 0 bytes")
            return render_template('index.html', error='Internal error: Generated PDF is empty')
        
        # Cleanup temp directory manually AFTER send_file to ensure file persists
        try:
            def cleanup():
                import atexit
                import signal
                import time
                
                time.sleep(1)  # Keep directory alive
                print("DEBUG: Cleanup executed")
                shutil.rmtree(tmpdir, ignore_errors=True)
            
            # Register cleanup to run on app exit
            atexit.register(cleanup)
        except Exception as e:
            print(f"ERROR: Could not register cleanup: {e}")
        
        print("DEBUG: Sending file...")
        
        return send_file(out_path, as_attachment=True, download_name='DD1750.pdf')
    
    except Exception as e:
        print(f"CRITICAL ERROR: {e}", file=sys.stderr)
        import traceback
        traceback.print_exc()
        return render_template('index.html', error=f"Error: {str(e)}")

if __name__ == '__main__':
    port = int(os.environ.get('PORT', 8000))
    app.run(host='0.0.0.0', port=port)
