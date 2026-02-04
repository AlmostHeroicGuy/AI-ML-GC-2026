import json
import time
from google import genai
from google.genai import types
from schema_guard import SectorGuard

class CostTracker:
    def __init__(self):
        self.total_cost_inr = 0
        self.session_cost = 0 
    def log(self, input_tokens, output_tokens):
        cost_usd = (input_tokens / 1e6 * 0.05) + (output_tokens / 1e6 * 0.20)
        self.total_cost_inr += (cost_usd * 84.0)
        self.session_cost += (cost_usd * 84.0)

class AnalysisAgent:
    def __init__(self, api_key):
        self.client = genai.Client(api_key=api_key)
        self.guard = SectorGuard()
        self.cost_tracker = CostTracker()
        self.active_model = None 
        self.PRIORITY_MODELS = ["gemini-2.0-flash-lite", "gemini-1.5-flash", "gemini-2.5-flash-lite"]
        self.SECTOR_DEFINITIONS = {
            "Pharma": [("pharmaceutical", 10), ("api", 10), ("drug", 10)],
            "Tech": [("saas", 10), ("software", 10)],
            "Manufacturing": [("plant", 10), ("factory", 10)],
            "Logistics": [("fleet", 10)],
            "Consumer Goods": [("fmcg", 10)],
            "D2C": [("ecommerce", 10)]
        }

    def test_api_connection(self):
        print("🔌 Negotiating Gemini...", end=" ")
        try:
            remotes = [m.name.replace("models/", "") for m in self.client.models.list()]
        except: remotes = []
        for c in self.PRIORITY_MODELS:
            matches = [m for m in remotes if c in m]
            if matches:
                try:
                    self.client.models.generate_content(model=matches[0], contents="Test")
                    self.active_model = matches[0]
                    print(f"✅ {self.active_model}")
                    return True
                except: continue
        
        fallback = "gemini-1.5-flash"
        self.active_model = fallback
        print(f"✅ Fallback: {fallback}")
        return True

    def _format_context_with_ids(self, chunks):
        MAX_CHARS = 1000000 
        context_str = "DATA VAULT (Cite these IDs):\n"
        sorted_chunks = sorted(chunks, key=lambda c: 3 if 'financial' in c['type'] else 1, reverse=True)
        for c in sorted_chunks:
            entry = f"[{c['id']}] SOURCE: {c['source']} ({c['location']})\n{c['text'][:40000]}\n\n"
            if len(context_str) + len(entry) > MAX_CHARS: break
            context_str += entry
        return context_str, set()

    def _sanitize(self, data, forbidden):
        if isinstance(data, dict): return {k: self._sanitize(v, forbidden) for k,v in data.items()}
        if isinstance(data, list): return [self._sanitize(i, forbidden) for i in data]
        if isinstance(data, str):
            clean = data.replace(forbidden, "Project X").replace(forbidden.replace(" ", ""), "Project X")
            if "source_display" in data and ("." in data or forbidden.lower() in data.lower()): return "Internal Doc"
            return clean
        return data

    def analyze_company(self, chunks, company_real_name):
        if not self.active_model: self.test_api_connection()
        print(f"🤖 Analyzing via {self.active_model}...")
        
        # Sector Heuristic
        scores = {k: 0 for k in self.SECTOR_DEFINITIONS}
        sample = " ".join([c['text'].lower() for c in chunks[:15]])
        for sec, kws in self.SECTOR_DEFINITIONS.items():
            for w, weight in kws:
                if w in sample: scores[sec] += (sample.count(w) * weight)
        best_sec = max(scores, key=scores.get)
        detected_sector = best_sec if scores[best_sec] > 5 else "General"
        print(f"🧠 Sector: {detected_sector}")

        context, _ = self._format_context_with_ids(chunks)
        
        prompt = f"""
        Strict M&A Analyst Task.
        INPUT: Name: "{company_real_name}" (FORBIDDEN). Sector: {detected_sector}.
        
        RULES:
        1. ANONYMIZE: Replace "{company_real_name}" with "Project X".
        2. CITATIONS: Use [ID]. In 'source_display', NEVER use filenames. Use "Internal Doc".
        3. FINANCIALS: Extract 'Revenue', 'EBITDA' for latest available year.
        4. OUTPUT JSON:
        {{
            "code_name": "Project X",
            "sector": "{detected_sector}",
            "slide_1": {{ "headline": "Key Investment Theme", "sub_headline": "One sentence summary", "bullets": ["Highlight 1", "Highlight 2 (Include Certifications)"] }},
            "slide_2": {{ "metrics": {{ "Revenue (Latest)": "100 Mn", "EBITDA": "20 Mn" }}, "chart_data": {{ "years": ["2022","2023","2024"], "revenue_values": [100, 120, 150], "data_quality": "Actuals" }} }},
            "slide_3": {{ "hooks": ["Strong Growth", "Market Leader", "High Margins", "Global Reach"] }},
            "citations": [ {{ "id": "...", "claim": "...", "source_display": "Internal Doc" }} ]
        }}
        """

        for attempt in range(3):
            try:
                print(f"⏳ Gen Attempt {attempt+1}...", end=" ", flush=True)
                resp = self.client.models.generate_content(
                    model=self.active_model, contents=[f"CONTEXT:\n{context}", prompt],
                    config=types.GenerateContentConfig(response_mime_type="application/json")
                )
                print("✅")
                
                res = json.loads(resp.text)
                res = self._sanitize(res, company_real_name)
                
                # Check Guardrails
                ok1, m1 = self.guard.check_anonymity(res, company_real_name)
                ok2, m2 = self.guard.validate(res)
                
                if ok1 and ok2: return res
                print(f"   ⚠️ Validation: {m1} | {m2}")
            
            except Exception as e:
                print(f"❌ {e}")
                time.sleep(2)
        return None