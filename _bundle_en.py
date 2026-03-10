"""Bundle English markdown files from root directory into _content.json."""
import json, os

DIR = os.path.dirname(os.path.abspath(__file__))

# Map section IDs to filenames in root
FILE_MAP = {
    "overview": "overview.md",
    "schedule": "schedule.md",
    "week-1": "week-1.md",
    "week-2": "week-2.md",
    "week-3": "week-3.md",
    "week-4": "week-4.md",
    "exercises": "exercises.md",
    "warmups": "warmups.md",
    "nutrition": "nutrition.md",
    "testing": "conditioning-tests.md",
    "sources": "sources.md",
    "safety": "safety.md",
}

data = {}
missing = []

for section_id, filename in FILE_MAP.items():
    path = os.path.join(DIR, filename)
    if os.path.exists(path):
        with open(path, encoding="utf-8") as f:
            data[section_id] = f.read()
        print(f"  {section_id}: {len(data[section_id]):,} chars from {filename}")
    else:
        missing.append(f"{section_id} ({filename})")

if missing:
    print(f"\nWARNING: Missing files: {', '.join(missing)}")

out = os.path.join(DIR, "_content.json")
with open(out, "w", encoding="utf-8") as f:
    json.dump(data, f, ensure_ascii=False)

size = os.path.getsize(out)
print(f"\nBundled _content.json: {size:,} bytes ({size/1024:.0f} KB)")
