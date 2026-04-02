#!/usr/bin/env python
"""PostToolUse hook for Read: immediately restore VB6 files back to native encoding after Read."""
import sys
import json
import os
import subprocess

sys.path.insert(0, os.path.dirname(__file__))
from vb6_config import get_encoding, is_vb6_file


def is_utf8(file_path):
    """Check if file is currently UTF-8 encoded."""
    try:
        result = subprocess.run(['file', file_path], capture_output=True, text=True)
        return 'UTF-8' in result.stdout
    except:
        return False


def main():
    data = json.load(sys.stdin)
    if data.get('tool_name') != 'Read':
        return

    file_path = data.get('tool_input', {}).get('file_path', '')
    if not file_path or not is_vb6_file(file_path) or not os.path.isfile(file_path):
        return

    # Only restore if the file is currently UTF-8 (was converted by PreToolUse)
    if not is_utf8(file_path):
        return

    enc = get_encoding()
    try:
        with open(file_path, 'r', encoding='utf-8') as f:
            text = f.read()

        text = text.replace('\r\n', '\n').replace('\n', '\r\n')

        with open(file_path, 'wb') as f:
            f.write(text.encode(enc))

    except Exception as e:
        print(json.dumps({
            "systemMessage": f"[vb6-hook] post-read restore failed: {e}"
        }))
        sys.exit(2)


if __name__ == '__main__':
    main()
