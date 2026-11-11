#!/usr/bin/env python3
"""Merge DEX files and runtime resources into one classes.jar."""
import glob
import os
import sys
import zipfile

dex_dir = sys.argv[1]
output = os.path.join(os.path.dirname(dex_dir), "classes.jar")
jars = sys.argv[2:]

seen = set()
with zipfile.ZipFile(output, "w", zipfile.ZIP_DEFLATED) as out:
    for dex in sorted(glob.glob(os.path.join(dex_dir, "classes*.dex"))):
        name = os.path.basename(dex)
        out.write(dex, name)
        seen.add(name)

    for jar in jars:
        if not os.path.isfile(jar):
            continue
        with zipfile.ZipFile(jar) as source:
            for info in source.infolist():
                if info.is_dir():
                    continue
                name = info.filename
                if name.endswith(".class") or name.endswith(".dex"):
                    continue
                if name in seen:
                    continue
                if name.startswith("META-INF/") and (
                    name.endswith(".SF") or name.endswith(".RSA") or name.endswith(".DSA")
                ):
                    continue
                seen.add(name)
                out.writestr(info, source.read(info))
EOF
mkdir -p /tmp/opencode/kino-office/scripts && mv /tmp/opencode/kino-office/scripts/package_resources.py /tmp/opencode/kino-office/scripts/package_resources.py 2>/dev/null; echo ok