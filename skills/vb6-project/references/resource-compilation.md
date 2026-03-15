# VB6 Resource File (.rc → .RES) Compilation

## Overview

VB6 projects use Windows resource files to store UI strings (STRINGTABLE), icons, bitmaps,
and other embedded resources. The workflow is:

1. Edit `.rc` source files (Big5 text)
2. Compile with `rc.exe` to produce `.RES` binary
3. VB6 links the `.RES` into the final `.exe` during compilation

The `.vbp` references the `.RES` file:
```
ResFile32="Source\Resource\CR_Bins200.RES"
```

## Directory Structure (typical)

```
Source/Resource/
├── CR_Bins200.rc          ← Main .rc file (entry point)
├── CR_Bins200.RES         ← Compiled binary (git tracked)
├── 12xx.rc                ← User interface strings (1200-1299)
├── 13xx.rc                ← Lot-related strings (1300-1399)
├── 14xx.rc                ← File-related strings
├── 15xx.rc                ← Error messages
├── ...                    ← More string ranges
├── 0xxx-Bins200IO.rc      ← I/O related strings
└── Others-Bins200.rc      ← Machine-specific strings
```

## .rc File Format

### STRINGTABLE (main format used)

The main `.rc` file wraps everything in a STRINGTABLE block and `#include`s sub-files:

```rc
STRINGTABLE
BEGIN
#include "12xx.rc"    //12xx - 使用者介面
#include "13xx.rc"    //13xx - 批次資訊
...
END
```

### Sub-file format (e.g., 12xx.rc)

Each sub-file contains string ID and value pairs:

```rc
//使用者介面 1200-1299
1200,"登入"
1201,"使用者帳號："
1202,"使用者密碼："
1203,"變更密碼"
```

- String IDs are numeric (no prefix)
- Values are double-quoted Big5 strings
- Comments use `//` prefix
- **Encoding: Big5/ANSI** — same rules as VB6 source files

### String ID Convention

| Range | Category |
|-------|----------|
| 0xxx | I/O descriptions |
| 12xx | User interface |
| 13xx | Lot information |
| 14xx | File operations |
| 15xx | Error messages |
| 16xx | Miscellaneous |
| 17xx | Motor-related strings |
| 18xx | Motor actions |
| 19xx | General control |
| 20xx | Screen interface |
| 21xx | Special actions |
| 22xx-25xx | Settings |
| 27xx | Communication |
| 29xx | Error definitions |
| 3000-4999 | Device descriptions (Others) |

### Accessing Strings in VB6 Code

Use `LoadResString()`:

```vb
Dim msg As String
msg = LoadResString(1200)  ' Returns "登入"
```

## Compiling .rc to .RES

### Using rc.exe from Windows SDK

```bash
RC_EXE="/c/Program Files (x86)/Windows Kits/10/bin/10.0.22621.0/x86/rc.exe"
"$RC_EXE" "C:\path\to\Source\Resource\CR_Bins200.rc"
```

This produces `CR_Bins200.RES` in the same directory.

### Important Notes

- **Use x86 rc.exe** — VB6 is 32-bit, use the x86 version of rc.exe for compatibility
- **Working directory matters** — `#include` paths in the .rc file are relative, so either
  cd to the Resource directory or use full paths
- **Encoding caution** — `.rc` files are Big5. Apply the same editing rules as other VB6 files
  (use `big5_edit.py` or python binary mode, never Edit/Write tools)
- **RES size may vary** — Different versions of rc.exe may produce slightly different .RES
  file sizes. The newer Windows SDK rc.exe may produce larger files than the original VB6-era
  compiler. Both should work with VB6, but verify by compiling the VB6 project after regenerating.
- **Always recompile VB6 after regenerating .RES** — The .RES is linked at VB6 compile time

### Bundled compile_rc.bat

The skill includes `scripts/compile_rc.bat` for convenient RC compilation:

```bash
COMPILE_RC="$HOME/.claude/plugins/local/vb6-dev/skills/vb6-project/scripts/compile_rc.bat"
cmd.exe //c "$COMPILE_RC \"C:\path\to\Source\Resource\CR_Bins200.rc\""
```

## Common Tasks

### Adding a New String

1. Determine the appropriate string ID range
2. Edit the corresponding sub-file (e.g., `15xx.rc` for error messages) using `big5_edit.py`
3. Recompile .RES
4. Use `LoadResString(ID)` in VB6 code

Example — adding error message 1560:

```bash
python ~/.claude/plugins/local/vb6-dev/skills/vb6-project/scripts/big5_edit.py \
  append "Source/Resource/15xx.rc" \
  '1560,"新的錯誤訊息"'
```

Then recompile .RES and the VB6 project.

### Modifying an Existing String

```bash
python ~/.claude/plugins/local/vb6-dev/skills/vb6-project/scripts/big5_edit.py \
  replace "Source/Resource/15xx.rc" \
  '1500,"舊的訊息"' \
  '1500,"新的訊息"'
```

### Adding a New String Range File

1. Create the new `.rc` file with `big5_write.py`
2. Add `#include "newfile.rc"` to `CR_Bins200.rc` using `big5_edit.py`
3. Recompile .RES

## Troubleshooting

| Issue | Cause | Fix |
|-------|-------|-----|
| `fatal error RC1107: invalid usage` | Wrong path format | Use Windows-style paths (`C:\...`) |
| Strings show garbled | .rc edited in UTF-8 | Restore from git, re-edit with Big5 tools |
| VB6 compile fails after .RES change | .RES format incompatible | Try using VB6 IDE's built-in resource editor instead |
| `LoadResString` returns wrong text | String ID collision | Check all .rc files for duplicate IDs |
