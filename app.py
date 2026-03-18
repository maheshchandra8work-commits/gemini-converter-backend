from flask import Flask, request, send_file
from flask_cors import CORS
import pypandoc
import os
import tempfile
import re

# This forces Render to download the core Pandoc software
pypandoc.download_pandoc()

app = Flask(__name__)
CORS(app)

@app.route('/')
def home():
    return "Omni-AI Word Conversion Service is running!"

@app.route('/convert-to-word', methods=['POST'])
def convert_to_word():
    data = request.json
    text = data.get('text', '')
    
    # --- THE OMNI-AI CLEANUP SCRIPT ---
    
    # 1. Rescue math formulas trapped inside hidden JSON widgets
    text = re.sub(r'genui[^]*"content":\s*"([^"]+)"[^]*', r'$$\n\1\n$$', text)
    
    # 2. Delete any leftover invisible tracking artifacts
    text = re.sub(r'[^]+', '', text)
    
    # 3. Fix broken Display Math from bad copy-pasting
    text = re.sub(r'^\[\s*$', '$$', text, flags=re.MULTILINE)
    text = re.sub(r'^\]\s*$', '$$', text, flags=re.MULTILINE)
    
    # 4. Fix standard Markdown Display Math
    text = text.replace(r'\[', '$$').replace(r'\]', '$$')
    
    # 5. Fix standard Markdown Inline Math
    text = text.replace(r'\(', '$').replace(r'\)', '$')

    # NEW: 6. Fix HTML alignment attributes so Word understands them
    # Word/Pandoc ignores raw align="center", but understands CSS text-align.
    text = re.sub(r'\balign="([^"]+)"', r'style="text-align: \1;"', text)
    
    with tempfile.NamedTemporaryFile(delete=False, suffix='.docx') as tmp:
        output_path = tmp.name
        
    try:
        # THE FIX: Two-Step Conversion Pipeline
        
        # Step A: Convert mixed Markdown + HTML into pure HTML (using MathML for equations)
        html_intermediate = pypandoc.convert_text(
            text, 
            'html', 
            format='markdown', 
            extra_args=['--mathml']
        )
        
        # Step B: Convert the resulting HTML directly into Word
        pypandoc.convert_text(
            html_intermediate, 
            'docx', 
            format='html', 
            outputfile=output_path
        )
        
        return send_file(output_path, as_attachment=True, download_name='AI_Document.docx')
    except Exception as e:
        print(f"Pandoc conversion error: {str(e)}")
        return {"error": "Failed to convert text"}, 500
    finally:
        if os.path.exists(output_path):
            os.remove(output_path)

if __name__ == '__main__':
    app.run(host='0.0.0.0', port=10000)
