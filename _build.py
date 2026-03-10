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

# Load marked.js for inlining
marked_path = os.path.join(DIR, "marked.min.js")
with open(marked_path, encoding="utf-8") as f:
    marked_js = f.read()

# Load local font CSS for inlining
fonts_css_path = os.path.join(DIR, "fonts", "fonts-local.css")
if os.path.exists(fonts_css_path):
    with open(fonts_css_path, encoding="utf-8") as f:
        fonts_css = f.read()
    fonts_tag = f"<style>{fonts_css}</style>"
else:
    print("WARNING: fonts/fonts-local.css not found, fonts will use system fallbacks")
    fonts_tag = ""

# Replace placeholders
html = template.replace("__CONTENT_EN_PLACEHOLDER__", en_json)
html = html.replace("__CONTENT_BOS_PLACEHOLDER__", bos_json)
html = html.replace("__MARKED_JS_PLACEHOLDER__", marked_js)
html = html.replace("__FONTS_CSS_PLACEHOLDER__", fonts_tag)

out = os.path.join(DIR, "index.html")
with open(out, "w", encoding="utf-8") as f:
    f.write(html)

size = os.path.getsize(out)
print(f"Built index.html: {size:,} bytes ({size/1024:.0f} KB)")
