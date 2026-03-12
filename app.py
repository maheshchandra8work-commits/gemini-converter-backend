from flask import Flask, request, send_file, jsonify
from flask_cors import CORS
import pypandoc
import os
import tempfile
import re
import json
import requests
from bs4 import BeautifulSoup

# This forces Render to download the core Pandoc software
pypandoc.download_pandoc()

app = Flask(__name__)
CORS(app)

@app.route('/')
def home():
    return "Your Render backend is awake and running!"

# --- THE AGGRESSIVE LINK SCRAPER ROUTE ---
@app.route('/scrape-link', methods=['POST'])
def scrape_link():
    data = request.get_json()
    url = data.get('url')

    if not url or ('chatgpt.com' not in url and 'gemini.google.com' not in url):
        return jsonify({"error": "Invalid URL. Please paste a ChatGPT or Gemini shared link."}), 400

    # Disguise our Python script as a real Google Chrome browser
    headers = {
        "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36",
        "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,image/avif,image/webp,*/*;q=0.8",
        "Accept-Language": "en-US,en;q=0.5"
    }

    try:
        res = requests.get(url, headers=headers, timeout=15)
        res.raise_for_status()
        html_content = res.text
        soup = BeautifulSoup(html_content, 'html.parser')
        extracted_text = ""

        # --- CHATGPT EXTRACTION ---
        if 'chatgpt.com' in url:
            # Look for the hidden JSON block where ChatGPT stores the message parts
            matches = re.findall(r'"parts":\s*(\[.*?\])', html_content)
            for m in matches:
                try:
                    parts = json.loads(m)
                    if parts and isinstance(parts[0], str):
                        extracted_text += parts[0] + "\n\n"
                except:
                    pass
            
            # Fallback if the JSON method fails
            if not extracted_text:
                for el in soup.find_all('div', class_=re.compile(r'markdown|prose')):
                    extracted_text += el.get_text(separator='\n') + "\n\n"

        # --- GEMINI EXTRACTION ---
        elif 'gemini.google.com' in url:
            # Gemini hides text in massive strings inside AF_initDataCallback arrays
            scripts = soup.find_all('script')
            for script in scripts:
                if script.string and 'AF_initDataCallback' in script.string:
                    # Look for extremely long text blocks that contain Markdown formatting
                    matches = re.findall(r'"([^"]{150,})"', script.string)
                    for m in matches:
                        clean_m = m.replace('\\n', '\n').replace('\\"', '"').replace('\\u003e', '>').replace('\\u003c', '<')
                        # Check if it looks like real text and not base64 computer code
                        if '\n' in clean_m and '{' not in clean_m[:10]:
                            extracted_text += clean_m + "\n\n"
                            break # Once we find the biggest text block, stop.

        extracted_text = extracted_text.strip()
        
        if not extracted_text:
            return jsonify({"error": "Could not extract text. The format may have changed, or the AI blocked the request."}), 404

        return jsonify({"text": extracted_text})

    except Exception as e:
        print("Scraping error:", str(e))
        return jsonify({"error": f"Server connection failed: {str(e)}"}), 500

# --- YOUR EXISTING WORD CONVERTER ROUTE ---
@app.route('/convert-to-word', methods=['POST'])
def convert_to_word():
    data = request.json
    text = data.get('text', '')
    
    # AI CLEANUP SCRIPT
    text = text.replace(r'\[', '$$').replace(r'\]', '$$')
    text = text.replace(r'\(', '$').replace(r'\)', '$')
    text = re.sub(r'genui[^]+', '', text)
    text = re.sub(r'(?m)^\\\[\s*$', '$$', text)
    text = re.sub(r'(?m)^\\\]\s*$', '$$', text)
    
    with tempfile.NamedTemporaryFile(delete=False, suffix='.docx') as tmp:
        output_path = tmp.name
        
    try:
        pypandoc.convert_text(text, 'docx', format='md', outputfile=output_path)
        return send_file(output_path, as_attachment=True, download_name='ai_document.docx')
    finally:
        if os.path.exists(output_path):
            os.remove(output_path)

if __name__ == '__main__':
    app.run(host='0.0.0.0', port=10000)
