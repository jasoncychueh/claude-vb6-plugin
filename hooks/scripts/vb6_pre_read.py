#!/usr/bin/env python
"""PreToolUse hook for Read: convert VB6 files from native encoding to UTF-8 before Read."""
import sys
import json
import os

sys.path.insert(0, os.path.dirname(__file__))
from vb6_config import get_encoding, is_vb6_file, is_native_encoded


def main():
    data = json.load(sys.stdin)
    if data.get('tool_name') != 'Read':
        return

    file_path = data.get('tool_input', {}).get('file_path', '')
    if not file_path or not is_vb6_file(file_path) or not os.path.isfile(file_path):
        return

    if not is_native_encoded(file_path):
        return

    enc = get_encoding()
    try:
        with open(file_path, 'rb') as f:
            raw = f.read()

        text = raw.decode(enc)

        with open(file_path, 'w', encoding='utf-8', newline='') as f:
            f.write(text)

        # Marker stores the encoding used, so restore knows what to convert back to
        marker = file_path + '.__vb6_encoding_pending__'
        with open(marker, 'w') as f:
            f.write(enc)

    except Exception as e:
        print(json.dumps({
            "systemMessage": f"[vb6-hook] {enc}→UTF-8 pre-read conversion failed: {e}"
        }))
        sys.exit(2)


if __name__ == '__main__':
    main()
