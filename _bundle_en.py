"""Bundle English markdown files from root directory into _content.json."""
import json, os

DIR = os.path.dirname(os.path.abspath(__file__))

# Map section IDs to filenames in root
FILE_MAP = {
    "dashboard": "dashboard.md",
    "phase-1": "phase-1.md",
    "phase-2": "phase-2.md",
    "phase-3": "phase-3.md",
    "phase-4": "phase-4.md",
    "weight-management": "weight-management.md",
    "training-tools": "training-tools.md",
    "fight-prep": "fight-prep.md",
    "schedule": "schedule.md",
    "sources": "sources.md",
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
