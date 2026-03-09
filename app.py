from flask import Flask, request, send_file
from flask_cors import CORS 
import pypandoc
import os
import tempfile

# This forces Render to download the core Pandoc software behind the scenes
pypandoc.download_pandoc()

app = Flask(__name__)
CORS(app) 

# A simple page to test if the server is awake
@app.route('/')
def home():
    return "Your Render backend is awake and running!"

@app.route('/convert-to-word', methods=['POST'])
def convert_to_word():
    data = request.json
    text = data.get('text', '')

    with tempfile.NamedTemporaryFile(delete=False, suffix='.docx') as tmp:
        output_path = tmp.name

    try:
        pypandoc.convert_text(text, 'docx', format='md', outputfile=output_path)
        return send_file(output_path, as_attachment=True, download_name='gemini_output.docx')
    finally:
        if os.path.exists(output_path):
            os.remove(output_path)

if __name__ == '__main__':
    app.run(host='0.0.0.0', port=10000)
  
