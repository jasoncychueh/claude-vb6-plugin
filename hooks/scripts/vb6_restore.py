#!/usr/bin/env python
"""Restore all pending files that were temporarily converted to UTF-8.

Scans for .__vb6_encoding_pending__ marker files and converts the corresponding
files back from UTF-8 to their native encoding + CRLF.

The marker file contains the encoding name (e.g., 'big5', 'shift_jis').

Can be called from:
- SessionEnd hook (Claude Code session ends)
- git pre-commit hook (before committing)
- compile.bat (before compiling)
- Command line directly: python big5_restore.py [search_dir]
"""
import os
import sys
import glob


def restore_file(file_path, marker_path):
    """Convert a single file from UTF-8 back to native encoding + CRLF."""
    try:
        # Read encoding from marker
        with open(marker_path, 'r') as f:
            enc = f.read().strip() or 'big5'

        with open(file_path, 'r', encoding='utf-8') as f:
            text = f.read()

        text = text.replace('\r\n', '\n').replace('\n', '\r\n')

        with open(file_path, 'wb') as f:
            f.write(text.encode(enc))

        os.remove(marker_path)
        return True
    except Exception as e:
        print(f'ERROR: Failed to restore {file_path}: {e}', file=sys.stderr)
        return False


def restore_all(search_dir='.'):
    """Find and restore all pending files."""
    pattern = os.path.join(search_dir, '**', '*.__vb6_encoding_pending__')
    markers = glob.glob(pattern, recursive=True)

    if not markers:
        return 0

    restored = 0
    for marker_path in markers:
        file_path = marker_path.replace('.__vb6_encoding_pending__', '')
        if os.path.isfile(file_path):
            if restore_file(file_path, marker_path):
                print(f'Restored: {file_path}')
                restored += 1
        else:
            os.remove(marker_path)

    return restored


def main():
    search_dir = sys.argv[1] if len(sys.argv) > 1 else '.'

    if not sys.stdin.isatty():
        try:
            sys.stdin.read()
        except:
            pass

    count = restore_all(search_dir)
    if count > 0:
        print(f'Restored {count} file(s)')


if __name__ == '__main__':
    main()
