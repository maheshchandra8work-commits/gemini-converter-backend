from flask import Flask, request, send_file
from flask_cors import CORS 
import pypandoc
import os
import tempfile
import re # We need this to find and fix weird AI code!

# This forces Render to download the core Pandoc software
pypandoc.download_pandoc()

app = Flask(__name__)
CORS(app) 

@app.route('/')
def home():
    return "Your Render backend is awake and running!"

@app.route('/convert-to-word', methods=['POST'])
def convert_to_word():
    data = request.json
    text = data.get('text', '')

    # --- THE ULTIMATE AI CLEANUP SCRIPT ---
    # 1. Fix standard ChatGPT brackets
    text = text.replace(r'\[', '$$').replace(r'\]', '$$')
    text = text.replace(r'\(', '$').replace(r'\)', '$')
    
    # 2. Fix Microsoft Copilot's weird "genui" math widgets
    # This finds that ugly JSON string and extracts just the math formula!
    text = re.sub(r'genui.*?\"content\":\s*\"(.*?)\".*?', r'$$\1$$', text)

    # 3. Fix broken brackets that lost their backslashes during a bad copy-paste
    # This turns a [ on its own line into a $$
    text = re.sub(r'(?m)^\[\s*$', '$$', text) 
    text = re.sub(r'(?m)^\]\s*$', '$$', text)
    # --------------------------------------

    with tempfile.NamedTemporaryFile(delete=False, suffix='.docx') as tmp:
        output_path = tmp.name

    try:
        pypandoc.convert_text(text, 'docx', format='md', outputfile=output_path)
        return send_file(output_path, as_attachment=True, download_name='ai_output.docx')
    finally:
        if os.path.exists(output_path):
            os.remove(output_path)

if __name__ == '__main__':
    app.run(host='0.0.0.0', port=10000)
