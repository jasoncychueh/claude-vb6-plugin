"""Shared config for vb6-project hooks. Reads encoding from .vb6-encoding file."""
import os
import subprocess

DEFAULT_ENCODING = 'big5'
CONFIG_FILENAME = '.vb6-encoding'

VB6_EXTENSIONS = {'.bas', '.cls', '.frm', '.rc'}


def get_project_root():
    """Get project root: CLAUDE_PROJECT_DIR > git root > cwd."""
    # Prefer Claude Code's project dir (always set in hooks)
    project_dir = os.environ.get('CLAUDE_PROJECT_DIR', '')
    if project_dir and os.path.isdir(project_dir):
        return project_dir
    try:
        result = subprocess.run(['git', 'rev-parse', '--show-toplevel'],
                                capture_output=True, text=True)
        root = result.stdout.strip()
        if root:
            return root
    except:
        pass
    return os.getcwd()


def get_encoding():
    """Read encoding from .vb6-encoding in project root. Default: big5."""
    root = get_project_root()
    config_path = os.path.join(root, CONFIG_FILENAME)
    if os.path.isfile(config_path):
        try:
            with open(config_path, 'r') as f:
                enc = f.read().strip().split('\n')[0].strip()
                if enc:
                    return enc
        except:
            pass
    return DEFAULT_ENCODING


def is_vb6_file(file_path):
    """Check if file has a VB6 extension."""
    ext = os.path.splitext(file_path)[1].lower()
    return ext in VB6_EXTENSIONS


def is_native_encoded(file_path):
    """Check if file is in native encoding (not UTF-8)."""
    try:
        result = subprocess.run(['file', file_path], capture_output=True, text=True)
        info = result.stdout
        return 'ISO-8859' in info or 'Non-ISO' in info
    except FileNotFoundError:
        # No 'file' command, try reading as the configured encoding
        try:
            enc = get_encoding()
            with open(file_path, 'rb') as f:
                f.read().decode(enc)
            return True
        except:
            return False


def create_config_if_missing():
    """Create .vb6-encoding with default encoding if it doesn't exist.
    Returns (created: bool, path: str)."""
    root = get_project_root()
    config_path = os.path.join(root, CONFIG_FILENAME)
    if os.path.isfile(config_path):
        return False, config_path

    with open(config_path, 'w') as f:
        f.write(DEFAULT_ENCODING + '\n')

    return True, config_path
