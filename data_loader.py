import os
import re
import pandas as pd
from pypdf import PdfReader
import requests
from bs4 import BeautifulSoup
import warnings
import unicodedata
import hashlib
import urllib3

urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)
warnings.filterwarnings('ignore', category=UserWarning, module='openpyxl')

try:
    from docx import Document
    HAS_DOCX = True
except ImportError:
    HAS_DOCX = False

class UniversalLoader:
    def __init__(self):
        self.headers = {
            'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/115.0.0.0 Safari/537.36'
        }

    def _generate_chunk_id(self, content, source, location):
        unique_str = f"{source}-{location}-{content[:20]}"
        return hashlib.md5(unique_str.encode()).hexdigest()[:8]

    def _clean_text(self, text):
        if not text: return ""
        text = unicodedata.normalize('NFKC', text)
        # RELAXED CLEANING: Keeps structure for Tables
        # Only removes control characters, keeps all punctuation and newlines
        text = re.sub(r'[\x00-\x08\x0b\x0c\x0e-\x1f\x7f]', '', text)
        # Collapse multiple spaces but PRESERVE NEWLINES for table rows
        text = re.sub(r'[ \t]+', ' ', text)
        return text.strip()

    def _read_excel(self, file_path):
        chunks = []
        filename = os.path.basename(file_path)
        try:
            xls = pd.ExcelFile(file_path)
            for sheet_name in xls.sheet_names:
                df = xls.parse(sheet_name)
                df = df.dropna(how='all').dropna(axis=1, how='all')
                text_content = df.to_markdown(index=False)
                if not text_content or len(text_content) < 10: continue

                lower_name = sheet_name.lower()
                chunk_type = "private_excel_generic"
                if any(x in lower_name for x in ['balance', 'p&l', 'profit', 'financial']):
                    chunk_type = "private_excel_financial"

                chunks.append({
                    "id": self._generate_chunk_id(text_content, filename, sheet_name),
                    "text": text_content,
                    "source": filename,
                    "location": f"Sheet: {sheet_name}",
                    "type": chunk_type
                })
        except Exception as e:
            print(f"❌ Error reading Excel {file_path}: {e}")
        return chunks

    def _read_pdf(self, file_path):
        chunks = []
        filename = os.path.basename(file_path)
        try:
            reader = PdfReader(file_path)
            for i, page in enumerate(reader.pages):
                text = self._clean_text(page.extract_text())
                if len(text) > 20:
                    loc = f"Page {i+1}"
                    chunks.append({
                        "id": self._generate_chunk_id(text, filename, loc),
                        "text": text,
                        "source": filename,
                        "location": loc,
                        "type": "private_pdf"
                    })
        except Exception as e:
            print(f"❌ Error reading PDF {file_path}: {e}")
        return chunks

    def _read_markdown(self, file_path):
        """Intelligent Markdown Splitter."""
        chunks = []
        filename = os.path.basename(file_path)
        try:
            with open(file_path, 'r', encoding='utf-8') as f:
                content = f.read()
            
            # Split by headers (Line starting with #)
            sections = re.split(r'(^|\n)#{1,3}\s+', content)
            
            current_chunk = ""
            current_header = "Intro"
            
            for part in sections:
                if not part.strip(): continue
                
                # Heuristic: If it's short and looks like a title, it's a header
                if len(part) < 100 and '\n' not in part.strip():
                    current_header = part.strip()
                    continue
                
                text = self._clean_text(part)
                if len(text) > 20:
                    # Determine type based on header
                    c_type = "private_text_generic"
                    h_low = current_header.lower()
                    if any(x in h_low for x in ['financial', 'revenue', 'profit', 'p&l', 'balance']):
                        c_type = "private_text_financial" # High priority!
                    elif any(x in h_low for x in ['about', 'profile', 'business']):
                        c_type = "private_text_about"

                    chunks.append({
                        "id": self._generate_chunk_id(text, filename, current_header),
                        "text": text,
                        "source": filename,
                        "location": f"Section: {current_header}",
                        "type": c_type
                    })
        except Exception as e:
            print(f"❌ Error reading Markdown: {e}")
            # Fallback
            return self._read_text_fallback(file_path)
        return chunks

    def _read_text_fallback(self, file_path):
        try:
            with open(file_path, 'r', encoding='utf-8') as f:
                t = self._clean_text(f.read())
                return [{
                    "id": self._generate_chunk_id(t, os.path.basename(file_path), "Full"),
                    "text": t,
                    "source": os.path.basename(file_path),
                    "location": "Full Text",
                    "type": "private_text"
                }]
        except: return []

    def _read_word(self, file_path):
        chunks = []
        if not HAS_DOCX: return chunks
        filename = os.path.basename(file_path)
        try:
            doc = Document(file_path)
            current_chunk = ""
            start_idx = 0
            for i, para in enumerate(doc.paragraphs):
                text = self._clean_text(para.text)
                if text:
                    current_chunk += text + "\n"
                if len(current_chunk) > 1000:
                    loc = f"Para {start_idx}-{i}"
                    chunks.append({
                        "id": self._generate_chunk_id(current_chunk, filename, loc),
                        "text": current_chunk,
                        "source": filename,
                        "location": loc,
                        "type": "private_docx"
                    })
                    current_chunk = ""
                    start_idx = i + 1
            if current_chunk:
                chunks.append({
                    "id": self._generate_chunk_id(current_chunk, filename, f"Para {start_idx}-End"),
                    "text": current_chunk,
                    "source": filename,
                    "location": f"Para {start_idx}-End",
                    "type": "private_docx"
                })
        except Exception as e:
            print(f"❌ Error reading Word: {e}")
        return chunks

    def _scrape_web(self, url):
        chunks = []
        try:
            print(f"🌐 Scraping: {url}")
            response = requests.get(url, headers=self.headers, timeout=15, verify=False)
            if response.status_code != 200:
                print(f"⚠️ Failed to scrape {url} ({response.status_code})")
                return chunks

            soup = BeautifulSoup(response.content, 'html.parser')
            for tag in soup(["script", "style", "nav", "footer", "iframe"]):
                tag.extract()

            headers = soup.find_all(['h1', 'h2', 'h3'])
            if headers:
                for header in headers:
                    section_title = self._clean_text(header.get_text())
                    content = ""
                    for sib in header.next_siblings:
                        if sib.name in ['h1', 'h2', 'h3']: break
                        if hasattr(sib, 'get_text'): content += sib.get_text(separator=" ", strip=True) + " "
                    
                    clean_content = self._clean_text(content)
                    if len(clean_content) > 50:
                        c_type = "public_web_general"
                        if 'about' in section_title.lower(): c_type = "public_web_about"
                        elif 'investor' in section_title.lower(): c_type = "public_web_financial"

                        chunks.append({
                            "id": self._generate_chunk_id(clean_content, url, section_title),
                            "text": clean_content[:4000],
                            "source": url,
                            "location": f"Section: {section_title[:50]}",
                            "type": c_type
                        })
            else:
                text = self._clean_text(soup.get_text(separator="\n"))
                chunks.append({
                    "id": self._generate_chunk_id(text, url, "Main"),
                    "text": text[:10000],
                    "source": url,
                    "location": "Main Page",
                    "type": "public_web_generic"
                })
        except Exception as e:
            print(f"❌ Scraping Error: {e}")
        return chunks

    def load_data(self, source):
        if source.startswith("http"): return self._scrape_web(source)
        if not os.path.exists(source): return []
        
        ext = source.split('.')[-1].lower()
        if ext in ['xlsx', 'xls']: return self._read_excel(source)
        if ext == 'pdf': return self._read_pdf(source)
        if ext in ['docx', 'doc']: return self._read_word(source)
        if ext in ['md', 'markdown']: return self._read_markdown(source) # SPECIFIC HANDLER
        
        # Fallback for .txt or unknown
        return self._read_text_fallback(source)