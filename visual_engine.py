import requests, random

class VisualEngine:
    def __init__(self, key):
        self.key = key
        self.headers = {"Authorization": key}
        self.audit_log = [] # FIXED: Restored
        
        self.vibes = {
            "Manufacturing": ["factory interior blur", "industrial automation"],
            "Pharma": ["laboratory research blur", "pharmaceutical production abstract"],
            "Logistics": ["warehouse blurred", "container ship aerial"],
            "Tech": ["abstract blue digital network", "server room bokeh"],
            "Consumer Goods": ["retail store blurred", "product packaging texture"],
            "General": ["modern office abstract", "business meeting blur"]
        }
        self.risky = ["logo", "text", "sign", "dashboard", "graph", "chart"]

    def fetch_image(self, kw, sector="General", slide_index=0):
        vibe = random.choice(self.vibes.get(sector, self.vibes["General"]))
        query = f"{vibe} {kw} no text"
        
        url = self._search(query)
        decision = "Smart Query Success"
        
        if not url:
            url = self._search(vibe)
            decision = "Fallback to Vibe"
            
        if url:
            self.audit_log.append({
                "slide": slide_index,
                "query": query,
                "url": url,
                "decision": decision
            })
            return url
        return None

    def _search(self, query):
        try:
            r = requests.get("https://api.pexels.com/v1/search", 
                             headers=self.headers, 
                             params={"query": query, "per_page": 10, "orientation": "landscape"}, 
                             timeout=5)
            if r.status_code == 200:
                photos = r.json().get('photos', [])
                for p in photos:
                    if not any(x in p.get('alt', '').lower() for x in self.risky):
                        return p['src']['large2x']
        except: pass
        return None

    def download_image(self, url, path):
        try:
            with open(path, 'wb') as f:
                f.write(requests.get(url, timeout=10).content)
            return True
        except: return False

    def get_audit_log(self): # FIXED: Restored
        return self.audit_log