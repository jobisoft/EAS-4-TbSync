#!/usr/bin/env python3
"""Pack the add-on into an installable .xpi.

Usage: python3 build.py [output-dir]      (default: dist)

Only the files of the add-on go into the package (manifest, background script,
content script, options page, icons). The helper app, the tests and this script
stay out. The packed add-on gets the time of the build in its name
("Outlook Web Helper (build 2026-11-08 04:12:05)") and in build.json, so that the
version that runs can be told apart from older ones. The source files are not
changed.
"""
import json
import os
import sys
import zipfile
from datetime import datetime

ROOT = os.path.dirname(os.path.abspath(__file__))
FILES = ["manifest.json", "background.js", "icon32.png", "icon64.png",
         "content/outlook-web.js",
         "options/options.html", "options/options.css", "options/options.js"]


def main():
    out_dir = sys.argv[1] if len(sys.argv) > 1 else os.path.join(ROOT, "dist")
    with open(os.path.join(ROOT, "manifest.json"), encoding="utf-8") as f:
        manifest = json.load(f)
    version = manifest["version"]
    stamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    os.makedirs(out_dir, exist_ok=True)
    xpi = os.path.join(out_dir, "outlook-web-helper_%s.xpi" % version.replace(".", "_"))
    if os.path.exists(xpi):
        os.remove(xpi)
    with zipfile.ZipFile(xpi, "w", zipfile.ZIP_DEFLATED) as z:
        for rel in FILES:
            if rel == "manifest.json":
                packed = dict(manifest, name="%s (build %s)" % (manifest["name"], stamp))
                z.writestr(rel, json.dumps(packed, indent=2) + "\n")
            else:
                z.write(os.path.join(ROOT, rel), rel)
        z.writestr("build.json", json.dumps({"version": version, "builtAt": stamp}) + "\n")
    print("Build time:", stamp)
    print("Created", xpi)


if __name__ == "__main__":
    main()
