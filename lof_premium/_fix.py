# -*- coding: utf-8 -*-
# Read _body.py, _body2.py, _body3.py, _word.py, _word2.py and append to build_v3_final.py
import os

extra_files = ['_body.py', '_body2.py', '_body3.py', '_word.py', '_word2.py']
existing = []
for fn in extra_files:
    if os.path.exists(fn):
        existing.append(fn)

print("Found extra files:", existing)

# Read current build_v3_final.py
with open('build_v3_final.py', 'r', encoding='utf-8') as f:
    base = f.read()

# Read and concatenate extras
extras = []
for fn in existing:
    with open(fn, 'r', encoding='utf-8') as f:
        content = f.read()
        # Remove the "# -*- coding" line if present
        lines = content.split('\n')
        lines = [l for l in lines if not l.startswith('# -*-')]
        extras.append('\n'.join(lines))

combined = base + '\n' + '\n'.join(extras)

with open('build_v3_final.py', 'w', encoding='utf-8') as f:
    f.write(combined)

print(f"Combined: {len(base)} + extras -> {len(combined)} chars")
