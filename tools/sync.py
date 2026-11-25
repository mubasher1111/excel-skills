"""Copy shared/xl_io.py into every skill that has a scripts/ folder, and build per-skill zips.

    python tools/sync.py          # copy shared helpers
    python tools/sync.py --zip    # also build dist/<skill>.zip for upload to Claude.ai
"""
import os
import shutil
import sys
import zipfile

LONG = "\\\\?\\"  # Windows prefix that allows paths over 260 characters
ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
if os.name == "nt" and not ROOT.startswith(LONG):
    ROOT = LONG + ROOT
SKILLS = os.path.join(ROOT, "plugins", "excel-skills", "skills")
SHARED = os.path.join(ROOT, "shared", "xl_io.py")

for name in sorted(os.listdir(SKILLS)):
    scripts = os.path.join(SKILLS, name, "scripts")
    if os.path.isdir(scripts):
        shutil.copyfile(SHARED, os.path.join(scripts, "xl_io.py"))
        print("synced", name)

if "--zip" in sys.argv:
    dist = os.path.join(ROOT, "dist")
    os.makedirs(dist, exist_ok=True)
    for name in sorted(os.listdir(SKILLS)):
        base = os.path.join(SKILLS, name)
        path = os.path.join(dist, f"{name}.zip")
        with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as z:
            for folder, dirs, files in os.walk(base):
                dirs[:] = [d for d in dirs if d != "__pycache__"]
                for f in files:
                    full = os.path.join(folder, f)
                    z.write(full, os.path.join(name, os.path.relpath(full, base)))
        print("zipped", path)
