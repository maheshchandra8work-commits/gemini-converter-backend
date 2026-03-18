from flask import Flask, request, send_file
from flask_cors import CORS
import pypandoc
import os
import tempfile
import re
from bs4 import BeautifulSoup

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

    with tempfile.NamedTemporaryFile(delete=False, suffix='.docx') as tmp:
        output_path = tmp.name
        
    try:
        # Step A: Convert Markdown to HTML
        html_intermediate = pypandoc.convert_text(
            text, 
            'html', 
            format='markdown', 
            extra_args=['--mathml']
        )
        
        # --- THE HTML ALIGNMENT INJECTOR ---
        soup = BeautifulSoup(html_intermediate, 'html.parser')
        
        # Catch 1: Find AI divs with <div align="center|right">
        for div in soup.find_all('div', align=True):
            alignment = div['align']
            # Push the alignment directly into the child paragraphs and headings
            for tag in div.find_all(['p', 'h1', 'h2', 'h3', 'h4', 'h5', 'h6']):
                existing_style = tag.get('style', '')
                tag['style'] = f"{existing_style} text-align: {alignment};".strip()
        
        # Catch 2: Find AI divs with <div style="text-align: center|right;">
        for div in soup.find_all('div', style=re.compile(r'text-align:\s*(left|center|right|justify)')):
            match = re.search(r'text-align:\s*(left|center|right|justify)', div['style'])
            if match:
                alignment = match.group(1)
                # Push the alignment directly into the child paragraphs and headings
                for tag in div.find_all(['p', 'h1', 'h2', 'h3', 'h4', 'h5', 'h6']):
                    existing_style = tag.get('style', '')
                    tag['style'] = f"{existing_style} text-align: {alignment};".strip()
        
        # Convert the modified Soup object back to a string
        html_intermediate = str(soup)
        
        # Step B: Convert the heavily corrected HTML directly into Word
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
