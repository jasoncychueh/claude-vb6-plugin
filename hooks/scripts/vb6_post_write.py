#!/usr/bin/env python
"""PostToolUse hook for Write: convert VB6 files from UTF-8 to native encoding + CRLF.

For .frm files, also auto-fixes common format issues:
- Begin lines missing trailing space
- Object lines missing proper quoting
"""
import sys
import json
import os
import re

sys.path.insert(0, os.path.dirname(__file__))
from vb6_config import get_encoding, is_vb6_file


def fix_frm_format(text):
    """Fix common .frm format issues that would cause VB6 load errors."""
    lines = text.split('\n')
    fixed = []
    for line in lines:
        stripped = line.rstrip('\r')
        # Begin lines must have trailing space
        if re.match(r'^(\s*)Begin\s+\S+\s+\S+$', stripped):
            stripped = stripped + ' '
        # Object line format: Object = "{GUID}#ver"; "name.OCX"
        m = re.match(r'^Object\s*=?\s*\{?([0-9A-Fa-f-]+)\}?#([^;]+);\s*"?([^"]+)"?\s*$', stripped)
        if m:
            guid, ver, ocx = m.group(1), m.group(2), m.group(3).strip()
            stripped = f'Object = "{{{guid}}}#{ver}"; "{ocx}"'
        fixed.append(stripped)
    return '\n'.join(fixed)


def main():
    data = json.load(sys.stdin)
    if data.get('tool_name') != 'Write':
        return

    file_path = data.get('tool_input', {}).get('file_path', '')
    if not file_path or not is_vb6_file(file_path) or not os.path.isfile(file_path):
        return

    enc = get_encoding()
    try:
        with open(file_path, 'r', encoding='utf-8') as f:
            text = f.read()

        # Auto-fix .frm format issues
        if os.path.splitext(file_path)[1].lower() == '.frm':
            text = fix_frm_format(text)

        text = text.replace('\r\n', '\n').replace('\n', '\r\n')

        with open(file_path, 'wb') as f:
            f.write(text.encode(enc))

        # Remove any pending marker (Write after Read scenario)
        marker = file_path + '.__vb6_encoding_pending__'
        if os.path.isfile(marker):
            os.remove(marker)

        msg = f"[vb6-hook] {file_path}: converted to {enc}+CRLF"
        print(json.dumps({"systemMessage": msg}))

    except Exception as e:
        print(json.dumps({
            "systemMessage": f"[vb6-hook] UTF-8→{enc} conversion failed for {file_path}: {e}"
        }))
        sys.exit(2)


if __name__ == '__main__':
    main()
