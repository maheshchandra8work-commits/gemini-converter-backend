from flask import Flask, request, send_file, jsonify
from flask_cors import CORS
import pypandoc
import os
import tempfile
import re
import requests
from bs4 import BeautifulSoup

# This forces Render to download the core Pandoc software
pypandoc.download_pandoc()

app = Flask(__name__)
CORS(app)

@app.route('/')
def home():
    return "Your Render backend is awake and running!"

# --- THE NEW LINK SCRAPER ROUTE ---
@app.route('/scrape-link', methods=['POST'])
def scrape_link():
    data = request.get_json()
    url = data.get('url')

    if not url or ('chatgpt.com' not in url and 'gemini.google.com' not in url):
        return jsonify({"error": "Invalid URL. Please paste a ChatGPT or Gemini shared link."}), 400

    headers = {
        "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36"
    }

    try:
        res = requests.get(url, headers=headers, timeout=10)
        res.raise_for_status()
        html_content = res.text
        soup = BeautifulSoup(html_content, 'html.parser')
        extracted_text = ""

        if 'chatgpt.com' in url:
            prose_elements = soup.find_all('div', class_=re.compile(r'prose'))
            if prose_elements:
                for el in prose_elements:
                    extracted_text += el.get_text(separator='\n') + "\n\n"
            else:
                scripts = soup.find_all('script')
                for script in scripts:
                    if script.string and 'markdown' in script.string:
                        matches = re.findall(r'"markdown":\s*"([^"]+)"', script.string)
                        for match in matches:
                            extracted_text += match.replace('\\n', '\n').replace('\\"', '"') + "\n\n"

        elif 'gemini.google.com' in url:
            elements = soup.find_all(['message-content', 'div'], class_=re.compile(r'model-response|response-container|message-content'))
            if elements:
                for el in elements:
                    extracted_text += el.get_text(separator='\n') + "\n\n"
            else:
                texts = re.findall(r'\["([^"]+)",null,\[', html_content)
                for text in texts:
                     extracted_text += text.replace('\\n', '\n') + "\n\n"

        if not extracted_text.strip():
            return jsonify({"error": "Could not find AI text in this link. The format may have changed."}), 404

        return jsonify({"text": extracted_text.strip()})

    except Exception as e:
        print("Scraping error:", str(e))
        return jsonify({"error": "Server error while fetching the link."}), 500

# --- YOUR EXISTING WORD CONVERTER ROUTE ---
@app.route('/convert-to-word', methods=['POST'])
def convert_to_word():
    data = request.json
    text = data.get('text', '')
    
    # --- THE ULTIMATE AI CLEANUP SCRIPT ---
    text = text.replace(r'\[', '$$').replace(r'\]', '$$')
    text = text.replace(r'\(', '$').replace(r'\)', '$')
    
    # Fix Microsoft Copilot's weird "genui" math widgets
    text = re.sub(r'genui[^]+', '', text)
    
    # Fix broken brackets that lost their backslashes during a bad copy-paste
    text = re.sub(r'(?m)^\\\[\s*$', '$$', text)
    text = re.sub(r'(?m)^\\\]\s*$', '$$', text)
    
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
