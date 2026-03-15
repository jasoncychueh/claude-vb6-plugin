---
name: vb6-project
description: >
  This skill should be used when editing VB6 source files (.bas, .cls, .frm, .frx),
  compiling VB6 projects (.vbp), creating new VB6 modules or forms, adding controls
  to forms, or troubleshooting VB6 compile errors. Also triggered by mentions of
  "Big5", "Shift-JIS", "GBK", "VB6 encoding", "compile VB6", ".frm format",
  "edit .bas", "new .frm", "add module", "add control", "CRLF", "encoding corruption",
  "resource string", ".rc file", ".gitattributes VB6", ".vb6-encoding".
  Provides encoding hooks for transparent Read/Edit/Write on ANSI-encoded VB6 files,
  .frm format auto-fix, compile helpers, and safety nets.
---

# VB6 Project Development

## How It Works — Encoding Protection

VB6 source files (.bas, .cls, .frm, .rc) are **ANSI encoded** (Big5, Shift-JIS, GBK, etc.) **with CRLF line endings**. Claude Code's Edit and Write tools output UTF-8, which would destroy these files.

This plugin uses a **hook-based transparent conversion** system to make Read/Edit/Write all safe:

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

1. **PreToolUse(Read)** — When a VB6 file is Read, the hook converts it from native encoding to UTF-8 on disk. A `.__vb6_encoding_pending__` marker file is created (containing the encoding name).
2. **Edit/Write runs normally** — The file is now UTF-8, so Edit/Write can work normally. Claude Code displays its native diff UI.
3. **PostToolUse(Edit/Write)** — After Edit or Write completes, the hook converts the file back to native encoding + CRLF and removes the marker.

### Encoding configuration

The encoding is read from `.vb6-encoding` in the project root (one line, e.g., `big5`). Supported values:

| Value | Language |
|-------|----------|
| `big5` | 繁體中文 (default) |
| `shift_jis` | 日本語 |
| `gbk` | 简体中文 |
| `euc-kr` | 한국어 |
| `cp1252` | Western European |

### Auto-setup on SessionStart

The plugin's **SessionStart hook** (`session_init.py`) automatically creates these files if they don't exist:

| File | Purpose |
|------|---------|
| `.vb6-encoding` | Encoding config (default: `big5`) |
| `.gitattributes` | CRLF line ending enforcement for VB6 files |
| `.git/hooks/pre-commit` | Safety net — restores pending files before commit |

No manual setup is needed.

### Safety nets for unconverted files

If a file is Read but never Edited (marker left behind), three layers of protection ensure it gets restored:

| Layer | Trigger | Mechanism |
|-------|---------|-----------|
| **SessionEnd hook** | Claude Code session ends | Scans for all `.__vb6_encoding_pending__` markers and restores |
| **git pre-commit hook** | `git commit` | Blocks commit if pending files found, restores them first |
| **compile.bat** | VB6 compilation | Restores all pending files before compiling |

**Important:** Do not run VB6 IDE while Claude Code is editing VB6 files. The temporary UTF-8 state on disk would corrupt VB6 IDE's view of the files.

## Critical Rules

### 1. Edit and Write tools are SAFE for VB6 files (hooks protect them)

The hooks automatically handle encoding conversion:
- **Edit** — Use normally for modifications. Claude Code displays native diff UI.
- **Write** — Use normally for creating new files. PostToolUse hook converts to native encoding + CRLF.
- **Read** — PreToolUse hook converts to UTF-8 so content is readable.

### 2. Forbidden tools (no hook protection)

| Tool | Problem |
|------|---------|
| `sed -i` | Strips CR from CRLF on Git Bash |
| `iconv` | Stops on first unconvertible char, truncates file |

## Creating New VB6 Files

**Use the Write tool directly.** The PostToolUse(Write) hook automatically converts to native encoding + CRLF.

### Creating `.frm` files — format auto-fixed by hook

`.frm` files have strict format requirements. The PostToolUse(Write) hook **automatically fixes** these common issues:

- **Begin trailing space** — `Begin VB.Form MyForm` → `Begin VB.Form MyForm ` (auto-added)
- **Object line quoting** — `Object={GUID}#ver; name.OCX` → `Object = "{GUID}#ver"; "name.OCX"` (auto-fixed)

You only need to get the **content structure** right (correct control nesting, property names, values).
The hook handles the format quirks. For full `.frm` format details, consult `references/frm-format.md`.

### Verify encoding

```bash
file Source/modified_file.bas
```

- **Correct:** `ISO-8859 text, with CRLF line terminators` (ANSI encodings show as ISO-8859)
- **Wrong:** `Unicode text, UTF-8 text, with CRLF line terminators`
- **Wrong:** `ISO-8859 text` (missing CRLF — line endings were stripped)

## Compiling

Use the project's own `compile.bat` if it has one, or the plugin's bundled version:

```bash
# If the project has its own compile.bat (preferred)
cmd.exe //c "cd /d <PROJECT_DIR> && <PROJECT_DIR>\compile.bat <PROJECT>.vbp"

# Or using the plugin's bundled compile.bat
cmd.exe //c "cd /d <PROJECT_DIR> && ${CLAUDE_PLUGIN_ROOT}\scripts\compile.bat <PROJECT>.vbp"
```

- `compile.bat` automatically restores any pending encoding files before compiling.
- VB6 command-line compilation shows OCX loading warnings (e.g., "無法載入 '0'") — these are **normal** and do not affect the build result.
- Check for `[OK] Compile succeeded.` at the end.
- If `[FAIL]`, read the `compile.log` and `.log` files next to the failing `.frm`.

### Compile error line numbers

VB6 compile errors report line numbers **relative to the code section** of `.frm` files, not the absolute file line.
For `.frm` files, the code section starts after the `Attribute VB_Exposed` line. To find the actual file line:

1. Find the `Attribute VB_Exposed` line number (e.g., line 163)
2. Add the reported error line number (e.g., 152)
3. Result = actual file line (e.g., 163 + 152 ≈ line 315)

For `.bas` and `.cls` files, the line number is counted from `Attribute VB_Name` (line 1 of the file), so it matches the file line directly.

## Resource Files (.rc → .RES)

VB6 uses `.rc` resource files (STRINGTABLE) for UI strings, compiled into `.RES` binary.
The `.rc` files are also ANSI encoded — Edit tool works safely on them (protected by hooks).

To recompile after editing `.rc` files:

```bash
cmd.exe //c "${CLAUDE_PLUGIN_ROOT}\scripts\compile_rc.bat \"<PROJECT_DIR>\Source\Resource\<main>.rc\""
```

For full details on string ID ranges, file structure, and troubleshooting,
consult `references/resource-compilation.md`.

## Registering New Files in .vbp

After creating new `.bas` or `.frm` files, add entries to the `.vbp` project file.
The `.vbp` is ASCII-safe, so the Edit tool is acceptable for this file.

- Module: `Module=ModuleName; Source\filename.bas`
- Form: `Form=Source\filename.frm`

## Additional Resources

All paths relative to `${CLAUDE_PLUGIN_ROOT}`:

### Scripts

- **`scripts/compile.bat`** — VB6 command-line compiler wrapper (auto-restores pending files)
- **`scripts/compile_rc.bat`** — Resource compiler wrapper (.rc → .RES)
- **`scripts/git-pre-commit`** — Git pre-commit hook (auto-installed by SessionStart)

### Hooks (auto-registered via hooks/hooks.json)

- **`hooks/scripts/session_init.py`** — SessionStart: creates `.vb6-encoding`, `.gitattributes`, installs git hook
- **`hooks/scripts/vb6_config.py`** — Shared config: reads `.vb6-encoding`, detects encoding
- **`hooks/scripts/vb6_pre_read.py`** — PreToolUse(Read): ANSI→UTF-8 transparent conversion
- **`hooks/scripts/vb6_post_edit.py`** — PostToolUse(Edit): UTF-8→ANSI+CRLF restore
- **`hooks/scripts/vb6_post_write.py`** — PostToolUse(Write): UTF-8→ANSI+CRLF + .frm format auto-fix
- **`hooks/scripts/vb6_restore.py`** — Shared restore logic for safety nets

### References

- **`references/frm-format.md`** — VB6 .frm file format specification and common mistakes
- **`references/resource-compilation.md`** — Resource file workflow, string ID ranges, troubleshooting
