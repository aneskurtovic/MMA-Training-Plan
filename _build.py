"""Build script: reads content JSON files and generates index.html with embedded content."""
import json, os

DIR = os.path.dirname(os.path.abspath(__file__))

# Load English content
with open(os.path.join(DIR, "_content.json"), encoding="utf-8") as f:
    en_data = json.load(f)

# Load Bosnian content
bos_path = os.path.join(DIR, "_content_bos.json")
if os.path.exists(bos_path):
    with open(bos_path, encoding="utf-8") as f:
        bos_data = json.load(f)
else:
    print("WARNING: _content_bos.json not found, using English as fallback for Bosnian")
    bos_data = en_data

en_json = json.dumps(en_data, ensure_ascii=False)
bos_json = json.dumps(bos_data, ensure_ascii=False)

# Read the HTML template
with open(os.path.join(DIR, "_template.html"), encoding="utf-8") as f:
    template = f.read()

# Replace placeholders
html = template.replace("__CONTENT_EN_PLACEHOLDER__", en_json)
html = html.replace("__CONTENT_BOS_PLACEHOLDER__", bos_json)

out = os.path.join(DIR, "index.html")
with open(out, "w", encoding="utf-8") as f:
    f.write(html)

size = os.path.getsize(out)
print(f"Built index.html: {size:,} bytes ({size/1024:.0f} KB)")
