import os
import tempfile
import subprocess
from flask import Flask, render_template, request, send_file
from dd1750_core import generate_dd1750_from_pdf

app = Flask(__name__)

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/generate', methods=['POST'])
def generate():
    if 'bom_file' not in request.files or 'template_file' not in request.files:
        return render_template('index.html')

    bom_file = request.files['bom_file']
    template_file = request.files['template_file']

    try:
        with tempfile.TemporaryDirectory() as tmpdir:
            bom_path = os.path.join(tmpdir, 'bom.pdf')
            template_path = os.path.join(tmpdir, 'template.pdf')
            items_path = os.path.join(tmpdir, 'items.pdf')
            combined_path = os.path.join(tmpdir, 'DD1750.pdf')

            bom_file.save(bom_path)
            template_file.save(template_path)

            # 1. Generate items PDF (simple, no template merge - this works perfectly)
            items_path, count = generate_dd1750_from_pdf(
                bom_pdf_path=bom_path,
                template_pdf_path=template_path,
                out_pdf_path=items_path
            )

            # 2. Check if user wants auto-combine
            auto_combine = request.form.get('auto_combine') == 'on'

            if auto_combine:
                # 2. Run pdftk to merge items with filled template
                # Command: pdftk A=items.pdf B=template.pdf output=combined.pdf
                try:
                    result = subprocess.run(
                        ['pdftk', 'A=items.pdf', 'B=template.pdf', 'output=combined.pdf'],
                        capture_output=True,
                        text=True,
                        check=True
                    )
                    
                    if result.returncode == 0:
                        # Success - return combined file
                        return send_file(combined_path, as_attachment=True, download_name='DD1750.pdf')
                    else:
                        # pdftk failed, return items file
                        return send_file(items_path, as_attachment=True, download_name='DD1750_items.pdf')
                
                except FileNotFoundError:
                    # pdftk not installed - return items file with instructions
                    print("pdftk not found - returning items only")
                    return send_file(items_path, as_attachment=True, download_name='DD1750_items.pdf')
                except Exception as e:
                    # Error - return items file
                    print(f"pdftk error: {e}")
                    return send_file(items_path, as_attachment=True, download_name='DD1750_items.pdf')
            
            else:
                # 3. No auto-combine - just return items PDF
                return send_file(items_path, as_attachment=True, download_name='DD1750_items.pdf')

    except Exception as e:
        return render_template('index.html', error=f"Error: {str(e)}")

if __name__ == '__main__':
    port = int(os.environ.get('PORT', 8000))
    app.run(host='0.0.0.0', port=port, debug=False)
