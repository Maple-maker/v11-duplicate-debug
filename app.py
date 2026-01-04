import os
import tempfile
from flask import Flask, render_template, request, send_file, flash, redirect
from dd1750_core import generate_dd1750_from_pdf

app = Flask(__name__)

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/generate', methods=['POST'])
def generate():
    if 'bom_file' not in request.files or 'template_file' not in request.files:
        flash('Both files required')
        return redirect('/')
    
    try:
        with tempfile.TemporaryDirectory() as tmp:
            bom_path = os.path.join(tmp, 'bom.pdf')
            tpl_path = os.path.join(tmp, 'tpl.pdf')
            out_path = os.path.join(tmp, 'out.pdf')
            
            request.files['bom_file'].save(bom_path)
            request.files['template_file'].save(tpl_path)
            
            out_path, count = generate_dd1750_from_pdf(bom_path, tpl_path, out_path)
            
            if count == 0:
                flash('No items found')
                return redirect('/')
            
            return send_file(out_path, as_attachment=True, download_name='DD1750.pdf')
    
    except Exception as e:
        flash(f'Error: {str(e)}')
        return redirect('/')

if __name__ == '__main__':
    app.run(host='0.0.0.0', port=int(os.environ.get('PORT', 8000)))
