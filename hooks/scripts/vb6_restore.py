#!/usr/bin/env python
"""Safety net: scan for VB6 files stuck in UTF-8 and restore to native encoding.

No longer relies on marker files. Scans Source/ for VB6 files that are UTF-8
and converts them back.

Can be called from:
- git pre-commit hook (before committing)
- compile.bat (before compiling)
- Command line directly: python vb6_restore.py [search_dir]
"""
import os
import sys
import glob
import subprocess

sys.path.insert(0, os.path.dirname(__file__))
from vb6_config import get_encoding, is_plugin_enabled, VB6_EXTENSIONS


def is_utf8(file_path):
    try:
        result = subprocess.run(['file', file_path], capture_output=True, text=True)
        return 'UTF-8' in result.stdout
    except:
        return False


def restore_all(search_dir='.'):
    """Find and restore all VB6 files stuck in UTF-8."""
    enc = get_encoding()
    restored = 0

    for ext in VB6_EXTENSIONS:
        pattern = os.path.join(search_dir, '**', '*' + ext)
        for file_path in glob.glob(pattern, recursive=True):
            if is_utf8(file_path):
                try:
                    with open(file_path, 'r', encoding='utf-8') as f:
                        text = f.read()
                    text = text.replace('\r\n', '\n').replace('\n', '\r\n')
                    with open(file_path, 'wb') as f:
                        f.write(text.encode(enc))
                    print(f'Restored: {file_path}')
                    restored += 1
                except Exception as e:
                    print(f'ERROR: {file_path}: {e}', file=sys.stderr)

    # Also clean up any leftover marker files
    for marker in glob.glob(os.path.join(search_dir, '**', '*.__vb6_encoding_pending__'), recursive=True):
        os.remove(marker)

    return restored


def main():
    search_dir = sys.argv[1] if len(sys.argv) > 1 else '.'

    if not sys.stdin.isatty():
        try:
            sys.stdin.read()
        except:
            pass
        # Called as Claude hook — check if plugin is enabled for this project
        if not is_plugin_enabled():
            return

    count = restore_all(search_dir)
    if count > 0:
        print(f'Restored {count} file(s)')


if __name__ == '__main__':
    main()
