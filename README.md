# claude-vb6-plugin

A [Claude Code](https://docs.anthropic.com/en/docs/claude-code) plugin that makes it safe to edit VB6 source files. VB6 uses ANSI encoding (Big5, Shift-JIS, GBK, etc.) with CRLF line endings — Claude Code's tools output UTF-8, which would corrupt these files. This plugin transparently handles the conversion.

## What it does

```
 Read VB6 file          Edit/Write VB6 file    After Edit/Write
 ┌─────────────┐       ┌──────────────┐       ┌──────────────┐
 │ PreToolUse   │       │ Tool runs    │       │ PostToolUse  │
 │ (Read hook)  │       │ normally on  │       │ (Edit/Write  │
 │              │       │ UTF-8 content│       │  hook)       │
 │ ANSI → UTF-8 │  ──>  │ Claude Code  │  ──>  │ UTF-8 → ANSI │
 │ on disk      │       │ shows native │       │ + CRLF       │
 │              │       │ diff UI      │       │ on disk      │
 └─────────────┘       └──────────────┘       └──────────────┘
```

- **Read** — Converts ANSI to UTF-8 before reading, so Claude can understand the content
- **Edit** — Works normally on UTF-8, then converts back. Native diff UI works as expected
- **Write** — Creates files in UTF-8, then converts. For `.frm` files, auto-fixes Begin trailing spaces and Object line formatting
- **Safety nets** — SessionEnd hook, git pre-commit hook, and compile.bat all restore files if conversion was interrupted

## Features

- Transparent ANSI ↔ UTF-8 conversion on Read/Edit/Write
- Multi-encoding support: Big5, Shift-JIS, GBK, EUC-KR, CP1252
- `.frm` format auto-fix (Begin trailing spaces, Object line quoting)
- Auto-creates `.gitattributes` for CRLF enforcement
- Auto-installs git pre-commit hook as safety net
- VB6 compile helpers (`compile.bat`, `compile_rc.bat`)

## Install

```bash
/plugin marketplace add jasoncychueh/claude-vb6-plugin
/plugin install vb6-project@claude-vb6-plugin
```

## Configuration

The plugin auto-creates `.vb6-encoding` in your project root on first session. Default encoding is `big5`. Change it to match your VB6 project's encoding:

```
big5          # 繁體中文 (default)
shift_jis     # 日本語
gbk           # 简体中文
euc-kr        # 한국어
cp1252        # Western European
```

## Auto-setup

On every Claude Code session start, the plugin automatically ensures:

| File | Purpose |
|------|---------|
| `.vb6-encoding` | Encoding configuration |
| `.gitattributes` | CRLF line ending enforcement |
| `.git/hooks/pre-commit` | Restores pending files before commit |

## Important

**Do not run VB6 IDE while Claude Code is editing VB6 files.** The plugin temporarily converts files to UTF-8 on disk during Read/Edit operations. VB6 IDE seeing UTF-8 content would cause corruption.

## License

MIT
