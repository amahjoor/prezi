import os
from flask import Flask, render_template, request, send_file, jsonify
from ppt_generator import PresentationGenerator
from dotenv import load_dotenv

# Load environment variables
load_dotenv()

app = Flask(__name__)
generator = PresentationGenerator()

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/generate', methods=['POST'])
def generate_presentation():
    prompt = request.form.get('prompt')
    convert_to_pdf = request.form.get('convert_to_pdf', 'true').lower() == 'true'
    
    if not prompt:
        return jsonify({'error': 'No prompt provided'}), 400
    
    try:
        # Generate presentation
        result = generator.generate(prompt, convert_to_pdf=convert_to_pdf)
        
        # Extract filenames from paths
        pptx_filename = os.path.basename(result['pptx_path'])
        pdf_filename = os.path.basename(result['pdf_path']) if result['pdf_path'] else None
        
        return jsonify({
            'success': True,
            'message': 'Presentation generated successfully',
            'pptx_filename': pptx_filename,
            'pdf_filename': pdf_filename,
            'has_pdf': pdf_filename is not None
        })
    except Exception as e:
        return jsonify({
            'success': False,
            'error': str(e)
        }), 500

@app.route('/download/<filename>')
def download_file(filename):
    file_path = os.path.join('generated', filename)
    return send_file(file_path, as_attachment=True)

@app.route('/api/generate', methods=['POST'])
def api_generate():
    data = request.json
    
    if not data or 'prompt' not in data:
        return jsonify({'error': 'No prompt provided in request'}), 400
    
    convert_to_pdf = data.get('convert_to_pdf', True)
    
    try:
        # Generate presentation
        result = generator.generate(data['prompt'], convert_to_pdf=convert_to_pdf)
        
        # Extract filenames from paths
        pptx_filename = os.path.basename(result['pptx_path'])
        pdf_filename = os.path.basename(result['pdf_path']) if result['pdf_path'] else None
        
        response_data = {
            'success': True,
            'message': 'Presentation generated successfully',
            'pptx_filename': pptx_filename,
            'pptx_download_url': f'/download/{pptx_filename}',
            'has_pdf': pdf_filename is not None
        }
        
        if pdf_filename:
            response_data['pdf_filename'] = pdf_filename
            response_data['pdf_download_url'] = f'/download/{pdf_filename}'
            
        return jsonify(response_data)
    except Exception as e:
        return jsonify({
            'success': False,
            'error': str(e)
        }), 500

if __name__ == '__main__':
    # Create generated directory if it doesn't exist
    if not os.path.exists('generated'):
        os.makedirs('generated')
        
    # Run the app
    app.run(debug=True) 