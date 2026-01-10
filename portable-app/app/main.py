"""
Report Card Generator - Web Application
Flask-based web UI for generating report cards.
"""

import os
import json
import uuid
import shutil
import zipfile
from pathlib import Path
from flask import Flask, render_template, request, jsonify, send_file, session
from werkzeug.utils import secure_filename

from services import ReportGenerator, load_settings, PDFConverter
from version import VERSION

# Initialize Flask app
app = Flask(__name__)
app.secret_key = os.urandom(24)

# Configuration
APP_DIR = Path(__file__).parent.parent
CONFIG_DIR = APP_DIR / "config"
UPLOAD_DIR = APP_DIR / "uploads"
OUTPUT_DIR = APP_DIR / "output"

# Ensure directories exist
UPLOAD_DIR.mkdir(parents=True, exist_ok=True)
OUTPUT_DIR.mkdir(parents=True, exist_ok=True)

# Load settings
SETTINGS_PATH = CONFIG_DIR / "settings.json"
settings = load_settings(str(SETTINGS_PATH)) if SETTINGS_PATH.exists() else {}


def allowed_file(filename: str, allowed_extensions: list) -> bool:
    """Check if file has allowed extension."""
    return '.' in filename and \
           filename.rsplit('.', 1)[1].lower() in [ext.lstrip('.') for ext in allowed_extensions]


@app.route('/')
def index():
    """Render the main page."""
    # Check LibreOffice availability
    try:
        pdf_converter = PDFConverter()
        libreoffice_available = pdf_converter.is_available()
        libreoffice_version = pdf_converter.get_version() if libreoffice_available else None
    except Exception:
        libreoffice_available = False
        libreoffice_version = None
    
    return render_template(
        'index.html',
        libreoffice_available=libreoffice_available,
        libreoffice_version=libreoffice_version,
        settings=settings,
        version=VERSION
    )


@app.route('/upload', methods=['POST'])
def upload_files():
    """Handle file uploads."""
    try:
        # Create unique session directory
        session_id = str(uuid.uuid4())
        session_dir = UPLOAD_DIR / session_id
        session_dir.mkdir(parents=True, exist_ok=True)
        
        uploaded_files = {}
        
        # Handle Excel file
        if 'excel_file' in request.files:
            excel_file = request.files['excel_file']
            if excel_file.filename and allowed_file(excel_file.filename, ['.xlsx', '.xls']):
                filename = secure_filename(excel_file.filename)
                excel_path = session_dir / filename
                excel_file.save(str(excel_path))
                uploaded_files['excel'] = str(excel_path)
        
        # Handle Word template
        if 'template_file' in request.files:
            template_file = request.files['template_file']
            if template_file.filename and allowed_file(template_file.filename, ['.docx']):
                filename = secure_filename(template_file.filename)
                template_path = session_dir / filename
                template_file.save(str(template_path))
                uploaded_files['template'] = str(template_path)
        
        # Handle mapping file
        if 'mapping_file' in request.files:
            mapping_file = request.files['mapping_file']
            if mapping_file.filename and allowed_file(mapping_file.filename, ['.json']):
                filename = secure_filename(mapping_file.filename)
                mapping_path = session_dir / filename
                mapping_file.save(str(mapping_path))
                uploaded_files['mapping'] = str(mapping_path)
        
        # Validate all files uploaded
        if len(uploaded_files) != 3:
            missing = []
            if 'excel' not in uploaded_files:
                missing.append('Excel file')
            if 'template' not in uploaded_files:
                missing.append('Word template')
            if 'mapping' not in uploaded_files:
                missing.append('Mapping JSON')
            return jsonify({
                'success': False,
                'error': f"Missing files: {', '.join(missing)}"
            }), 400
        
        return jsonify({
            'success': True,
            'session_id': session_id,
            'files': uploaded_files
        })
    
    except Exception as e:
        return jsonify({
            'success': False,
            'error': str(e)
        }), 500


@app.route('/generate', methods=['POST'])
def generate_reports():
    """Generate report cards."""
    try:
        data = request.get_json()
        
        session_id = data.get('session_id')
        class_name = data.get('class_name', '').strip()
        excel_path = data.get('excel_path')
        template_path = data.get('template_path')
        mapping_path = data.get('mapping_path')
        
        # Validate inputs
        if not all([session_id, class_name, excel_path, template_path, mapping_path]):
            return jsonify({
                'success': False,
                'error': 'Missing required parameters'
            }), 400
        
        # Create output directory for this session
        output_dir = OUTPUT_DIR / session_id / f"{class_name}_report_cards"
        output_dir.mkdir(parents=True, exist_ok=True)
        
        # Initialize generator with settings
        generator = ReportGenerator(settings=settings)
        
        # Validate inputs first
        validation_errors = generator.validate_inputs(excel_path, template_path, mapping_path)
        if validation_errors:
            return jsonify({
                'success': False,
                'error': 'Validation failed',
                'errors': validation_errors
            }), 400
        
        # Generate reports
        results = generator.generate(
            excel_path=excel_path,
            template_path=template_path,
            mapping_path=mapping_path,
            class_name=class_name,
            output_dir=str(output_dir)
        )
        
        # Create ZIP file if successful
        if results['generated'] > 0:
            zip_path = OUTPUT_DIR / session_id / f"{class_name}_report_cards.zip"
            with zipfile.ZipFile(str(zip_path), 'w', zipfile.ZIP_DEFLATED) as zipf:
                for pdf_file in output_dir.glob('*.pdf'):
                    zipf.write(pdf_file, pdf_file.name)
            results['zip_path'] = str(zip_path)
            results['download_url'] = f"/download/{session_id}/{class_name}_report_cards.zip"
        
        return jsonify(results)
    
    except Exception as e:
        return jsonify({
            'success': False,
            'error': str(e)
        }), 500


@app.route('/download/<session_id>/<filename>')
def download_file(session_id: str, filename: str):
    """Download generated file."""
    try:
        file_path = OUTPUT_DIR / session_id / secure_filename(filename)
        
        if not file_path.exists():
            return jsonify({'error': 'File not found'}), 404
        
        return send_file(
            str(file_path),
            as_attachment=True,
            download_name=filename
        )
    
    except Exception as e:
        return jsonify({'error': str(e)}), 500


@app.route('/cleanup', methods=['POST'])
def cleanup():
    """Clean up session files."""
    try:
        data = request.get_json()
        session_id = data.get('session_id')
        
        if session_id:
            # Clean upload directory
            upload_session = UPLOAD_DIR / session_id
            if upload_session.exists():
                shutil.rmtree(upload_session)
            
            # Clean output directory
            output_session = OUTPUT_DIR / session_id
            if output_session.exists():
                shutil.rmtree(output_session)
        
        return jsonify({'success': True})
    
    except Exception as e:
        return jsonify({'success': False, 'error': str(e)}), 500


@app.route('/health')
def health():
    """Health check endpoint."""
    try:
        pdf_converter = PDFConverter()
        libreoffice_ok = pdf_converter.is_available()
    except Exception:
        libreoffice_ok = False
    
    return jsonify({
        'status': 'ok' if libreoffice_ok else 'degraded',
        'libreoffice': libreoffice_ok
    })


if __name__ == '__main__':
    import webbrowser
    import threading
    
    port = int(os.environ.get('PORT', 8080))
    url = f"http://localhost:{port}"
    
    # Open browser after a short delay
    def open_browser():
        import time
        time.sleep(1)
        webbrowser.open(url)
    
    threading.Thread(target=open_browser, daemon=True).start()
    
    print(f"\n{'='*50}")
    print("  Report Card Generator")
    print(f"  Running at: {url}")
    print(f"{'='*50}\n")
    
    app.run(host='0.0.0.0', port=port, debug=False)

