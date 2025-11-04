from flask import Flask, render_template, request, send_file, jsonify, after_this_request
import os
import uuid
import threading
import time
from pathlib import Path
from werkzeug.utils import secure_filename
from converter import convert_pdf_to_ppt

app = Flask(__name__)

# Configuration
UPLOAD_FOLDER = Path('temp_uploads')
OUTPUT_FOLDER = Path('temp_outputs')
MAX_FILE_SIZE = 50 * 1024 * 1024  # 50 MB
ALLOWED_EXTENSIONS = {'pdf'}
CLEANUP_DELAY = 300  # 5 minutes in seconds

# Create folders if they don't exist
UPLOAD_FOLDER.mkdir(exist_ok=True)
OUTPUT_FOLDER.mkdir(exist_ok=True)

def allowed_file(filename):
    """Check if file has allowed extension."""
    return '.' in filename and \
           filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS

def schedule_file_deletion(file_paths, delay=CLEANUP_DELAY):
    """
    Schedule files for deletion after a delay.
    
    Args:
        file_paths: List of file paths to delete
        delay: Delay in seconds before deletion (default: 300 = 5 minutes)
    """
    def delete_files():
        time.sleep(delay)
        for file_path in file_paths:
            try:
                if os.path.exists(file_path):
                    os.remove(file_path)
                    print(f"Deleted file: {file_path}")
            except Exception as e:
                print(f"Error deleting file {file_path}: {e}")
    
    # Run deletion in background thread
    thread = threading.Thread(target=delete_files, daemon=True)
    thread.start()

def cleanup_old_files(folder, max_age_seconds=3600):
    """
    Clean up old files in a folder (files older than max_age_seconds).
    
    Args:
        folder: Folder path to clean
        max_age_seconds: Maximum age of files in seconds (default: 1 hour)
    """
    try:
        current_time = time.time()
        for file_path in Path(folder).glob('*'):
            if file_path.is_file():
                file_age = current_time - file_path.stat().st_mtime
                if file_age > max_age_seconds:
                    try:
                        file_path.unlink()
                        print(f"Cleaned up old file: {file_path}")
                    except Exception as e:
                        print(f"Error cleaning up {file_path}: {e}")
    except Exception as e:
        print(f"Error during cleanup: {e}")

def immediate_file_cleanup(file_paths):
    """
    Immediately delete files after response is sent.
    
    Args:
        file_paths: List of file paths to delete
    """
    for file_path in file_paths:
        try:
            if os.path.exists(file_path):
                os.remove(file_path)
                print(f"Immediately deleted: {file_path}")
        except Exception as e:
            print(f"Error in immediate cleanup of {file_path}: {e}")

@app.route('/')
def index():
    """Render the main page."""
    # Clean up old files on page load
    cleanup_old_files(UPLOAD_FOLDER)
    cleanup_old_files(OUTPUT_FOLDER)
    return render_template('index.html')

@app.route('/convert', methods=['POST'])
def convert():
    """Handle PDF to PPT conversion."""
    
    # Validate request
    if 'pdf' not in request.files:
        return jsonify({'error': 'No file provided'}), 400
    
    file = request.files['pdf']
    
    if file.filename == '':
        return jsonify({'error': 'No file selected'}), 400
    
    if not allowed_file(file.filename):
        return jsonify({'error': 'Invalid file type. Only PDF files are allowed'}), 400
    
    # Generate unique filenames to prevent collisions
    unique_id = str(uuid.uuid4())
    secure_name = secure_filename(file.filename)
    file_path = UPLOAD_FOLDER / f"{unique_id}_{secure_name}"
    output_filename = f"{unique_id}_converted.pptx"
    output_path = OUTPUT_FOLDER / output_filename
    
    try:
        # Save uploaded file
        file.save(str(file_path))
        
        # Check file size after saving
        if file_path.stat().st_size > MAX_FILE_SIZE:
            file_path.unlink()  # Delete the file
            return jsonify({'error': 'File too large. Maximum size is 50MB'}), 413
        
        # Convert PDF to PPT
        try:
            convert_pdf_to_ppt(str(file_path), str(output_path))
        except Exception as e:
            # Clean up on conversion error
            if file_path.exists():
                file_path.unlink()
            return jsonify({'error': f'Conversion failed: {str(e)}'}), 500
        
        # Verify output was created
        if not output_path.exists():
            if file_path.exists():
                file_path.unlink()
            return jsonify({'error': 'Conversion failed to produce output'}), 500
        
        # Register cleanup after response is sent (immediate cleanup)
        @after_this_request
        def cleanup_after_request(response):
            """Delete files immediately after sending the response."""
            try:
                # Small delay to ensure file is fully sent
                threading.Timer(2.0, immediate_file_cleanup, 
                              args=([str(file_path), str(output_path)],)).start()
            except Exception as e:
                print(f"Error scheduling immediate cleanup: {e}")
            return response
        
        # Also schedule delayed cleanup as fallback (in case immediate cleanup fails)
        schedule_file_deletion([str(file_path), str(output_path)], delay=CLEANUP_DELAY)
        
        # Send file with original name (without UUID)
        original_name = Path(secure_name).stem + '.pptx'
        
        return send_file(
            str(output_path),
            as_attachment=True,
            download_name=original_name,
            mimetype='application/vnd.openxmlformats-officedocument.presentationml.presentation'
        )
    
    except Exception as e:
        # Clean up on any error
        if file_path.exists():
            try:
                file_path.unlink()
            except:
                pass
        if output_path.exists():
            try:
                output_path.unlink()
            except:
                pass
        
        return jsonify({'error': f'Server error: {str(e)}'}), 500

@app.route('/health')
def health():
    """Health check endpoint."""
    return jsonify({'status': 'healthy'}), 200

@app.errorhandler(413)
def request_entity_too_large(error):
    """Handle file too large error."""
    return jsonify({'error': 'File too large'}), 413

@app.errorhandler(500)
def internal_server_error(error):
    """Handle internal server errors."""
    return jsonify({'error': 'Internal server error'}), 500

if __name__ == "__main__":
    # Set max content length (50 MB)
    app.config['MAX_CONTENT_LENGTH'] = MAX_FILE_SIZE
    
    # Run with debug=False in production
    app.run(debug=True, host='0.0.0.0', port=5000)