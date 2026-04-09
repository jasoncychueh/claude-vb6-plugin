#!/usr/bin/env python
"""SessionStart hook: auto-install git pre-commit hook and create .vb6-encoding config."""
import sys
import json
import os
import shutil
import subprocess

sys.path.insert(0, os.path.dirname(__file__))
from vb6_config import get_encoding, is_plugin_enabled


def main():
    try:
        sys.stdin.read()
    except:
        pass

    if not is_plugin_enabled():
        return

    messages = []

    # Find project root
    project_dir = os.environ.get('CLAUDE_PROJECT_DIR', '')
    if not project_dir or not os.path.isdir(project_dir):
        try:
            result = subprocess.run(['git', 'rev-parse', '--show-toplevel'],
                                    capture_output=True, text=True)
            project_dir = result.stdout.strip()
        except:
            pass
    if not project_dir:
        project_dir = os.getcwd()

    # 1. Create .vb6-encoding if missing
    config_path = os.path.join(project_dir, '.vb6-encoding')
    if not os.path.isfile(config_path):
        try:
            with open(config_path, 'w') as f:
                f.write('big5\n')
            messages.append("Created .vb6-encoding")
        except Exception as e:
            messages.append(f".vb6-encoding failed: {e}")

    # 2. Create .gitattributes if missing
    gitattr_path = os.path.join(project_dir, '.gitattributes')
    if not os.path.isfile(gitattr_path):
        try:
            with open(gitattr_path, 'w') as f:
                f.write(
                    '*.bas eol=crlf\n'
                    '*.cls eol=crlf\n'
                    '*.frm eol=crlf\n'
                    '*.frx binary\n'
                    '*.vbp eol=crlf\n'
                    '*.vbw eol=crlf\n'
                    '*.rc  eol=crlf\n'
                    '*.RES binary\n'
                    '*.bat eol=crlf\n'
                    '*.dll binary\n'
                    '*.exe binary\n'
                    '*.ocx binary\n'
                )
            messages.append("Created .gitattributes")
        except Exception as e:
            messages.append(f".gitattributes failed: {e}")

    # 3. Install git pre-commit hook
    git_dir = os.path.join(project_dir, '.git')
    if os.path.isdir(git_dir):
        hook_target = os.path.join(git_dir, 'hooks', 'pre-commit')
        if not os.path.isfile(hook_target):
            plugin_root = os.environ.get('CLAUDE_PLUGIN_ROOT', '')
            if plugin_root:
                hook_source = os.path.join(plugin_root, 'scripts', 'git-pre-commit')
            else:
                hook_source = os.path.join(os.path.dirname(os.path.abspath(__file__)),
                                           '..', '..', 'scripts', 'git-pre-commit')
                hook_source = os.path.normpath(hook_source)

            if os.path.isfile(hook_source):
                os.makedirs(os.path.dirname(hook_target), exist_ok=True)
                shutil.copy2(hook_source, hook_target)
                try:
                    os.chmod(hook_target, 0o755)
                except:
                    pass
                messages.append("Installed git pre-commit hook")

    if not messages:
        messages.append("Ready (encoding: " + get_encoding() + ")")

    print(json.dumps({
        "systemMessage": "[vb6-plugin] " + "; ".join(messages)
    }))


if __name__ == '__main__':
    main()
