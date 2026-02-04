import os
import sys
from google import genai

# Get Key
key = os.getenv("GEMINI_API_KEY")
if not key:
    print("❌ Error: GEMINI_API_KEY not found.")
    sys.exit(1)

print(f"🔑 Using Key: {key[:5]}...")
print("d📡 Connecting to Google AI...")

try:
    client = genai.Client(api_key=key)
    
    print("\n📋 AVAILABLE MODELS:")
    print("===================")
    
    # New SDK syntax for listing models
    for m in client.models.list():
        # We only care about models that support generation
        if "generateContent" in m.supported_generation_methods:
            print(f"✅ {m.name}")
            
except Exception as e:
    print(f"\n❌ CRITICAL ERROR: {str(e)}")