#!/usr/bin/env python
"""PostToolUse hook for Edit: convert VB6 files back from native encoding + CRLF after Edit."""
import sys
import json
import os


def main():
    data = json.load(sys.stdin)
    if data.get('tool_name') != 'Edit':
        return

    file_path = data.get('tool_input', {}).get('file_path', '')
    if not file_path:
        return

    # Check for the marker left by PreToolUse(Read)
    marker = file_path + '.__vb6_encoding_pending__'
    if not os.path.isfile(marker):
        return

    # Read encoding from marker (stored by pre_read hook)
    with open(marker, 'r') as f:
        enc = f.read().strip() or 'big5'
    os.remove(marker)

    if not os.path.isfile(file_path):
        return

    try:
        with open(file_path, 'r', encoding='utf-8') as f:
            text = f.read()

        text = text.replace('\r\n', '\n').replace('\n', '\r\n')

        with open(file_path, 'wb') as f:
            f.write(text.encode(enc))

        msg = f"[vb6-hook] {file_path}: {enc}+CRLF restored"
        print(json.dumps({"systemMessage": msg}))

    except Exception as e:
        print(json.dumps({
            "systemMessage": f"[vb6-hook] UTF-8→{enc} conversion failed: {e}. File may be corrupted!"
        }))
        sys.exit(2)


if __name__ == '__main__':
    main()
