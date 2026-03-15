# VB6 .frm File Format Reference

## Overview

VB6 `.frm` files are **plain text** (Big5/ANSI encoded, CRLF line endings) with a strict
format. The file has two sections: the **form definition** (visual layout) and the
**code section** (VB6 source code).

## File Structure

```
VERSION 5.00
[Object references]
Begin VB.Form FormName ←── trailing space required!
   [Form properties]
   Begin VB.ControlType ControlName ←── trailing space required!
      [Control properties]
   End
   [More controls...]
End
Attribute VB_Name = "FormName"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
[Option Explicit]
[Variable declarations]
[Sub/Function code]
```

## Critical Format Rules

### 1. Begin Lines Must Have Trailing Space

Every `Begin` line **must** end with a space character before CRLF:

```
Begin VB.Form brmMyForm \r\n    ← space before \r\n
   Begin VB.CommandButton CmdOK \r\n    ← space before \r\n
```

**Without the trailing space, VB6 cannot load the form.**

### 2. Object Reference Format

OCX references must follow this exact format with spaces and quotes:

```
Object = "{GUID}#version"; "FILENAME.OCX"
```

**Correct:**
```
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
```

**Wrong — all of these will fail:**
```
Object={5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0; MSFLXGRD.OCX     ← missing spaces and quotes
Object = {5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0; "MSFLXGRD.OCX"  ← missing GUID quotes
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; MSFLXGRD.OCX  ← missing OCX quotes
```

### 3. Common OCX GUIDs

| OCX | GUID | Version |
|-----|------|---------|
| MSFLXGRD.OCX | {5E9E78A0-531B-11CF-91F6-C2863C385E30} | #1.0#0 |
| TABCTL32.OCX | {BDC217C8-ED16-11CD-956C-0000C04E4C0A} | #1.1#0 |
| MSWINSCK.OCX | {248DD890-BB45-11CF-9ABC-0080C7E7B78D} | #1.0#0 |
| COMCTL32.OCX | {6B7E6392-850A-101B-AFC0-4210102A8DA7} | #2.1#0 |
| RICHTX32.OCX | {3B7C8863-D78F-101B-B9B5-04021C009402} | #1.2#0 |
| COMDLG32.OCX | {F9043C88-F6F2-101A-A3C9-08002B2F49FB} | #1.2#0 |
| COMCT232.OCX | {86CF1D34-0C5F-11D2-A9FC-0000F8754DA1} | #2.0#0 |

### 4. Property Alignment

Properties use fixed-width alignment with `=` padded:

```
      Caption         =   "OK"
      Height          =   315
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1200
```

- Property name: left-aligned, padded to 16 chars
- `=` preceded by spaces to align
- Value preceded by 3 spaces

### 5. Indentation

- Form-level properties: 3 spaces
- Control Begin/End: 3 spaces
- Control properties: 6 spaces
- Nested control Begin/End: 6 spaces
- Nested control properties: 9 spaces

### 6. BeginProperty / EndProperty

Font and other compound properties use BeginProperty blocks:

```
   BeginProperty Font
      Name            =   "新細明體"
      Size            =   9.75
      Charset         =   136
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
```

Note: `BeginProperty` also requires trailing space.

### 7. Boolean Values

VB6 .frm uses these boolean formats:

```
      ControlBox      =   0   'False
      MaxButton       =   0   'False
      Cancel          =   -1  'True
```

`0` = False, `-1` = True, with the comment after.

### 8. Color Values

Hex color format with `&H` prefix and `&` suffix:

```
      BackColor       =   &H80000005&
      ForeColor       =   &H00FF0000&
```

## Adding Controls to Existing Forms

To add a new control to an existing form:

1. Read the file in binary mode
2. Find the last `   End\r\nEnd\r\n` boundary (last control End + form End)
3. Insert the new control block before `End\r\n` (form End)
4. Ensure all Begin lines have trailing space
5. Ensure CRLF line endings throughout

### Example: Adding a CommandButton

```python
btn = (
    b'   Begin VB.CommandButton CmdMyButton \r\n'
    b'      Caption         =   "My Button"\r\n'
    b'      Height          =   420\r\n'
    b'      Left            =   120\r\n'
    b'      TabIndex        =   38\r\n'
    b'      Top             =   4680\r\n'
    b'      Width           =   2895\r\n'
    b'   End\r\n'
)
```

## Adding Event Handlers

Event handler code goes at the end of the file, after all existing `End Sub` blocks:

```python
evt = (
    b'\r\n'
    b'Private Sub CmdMyButton_Click()\r\n'
    b'  ' + b"'" + '中文註解'.encode('big5') + b'\r\n'
    b'  MsgBox "Hello"\r\n'
    b'End Sub\r\n'
)
# Append to file data
data = data + evt
```

## Common Errors and Causes

| Error Message | Cause |
|--------------|-------|
| 無法載入 '0' | OCX control can't load in CLI mode (usually harmless) |
| 前置資料已被開啟 | .frm format error — check Object lines and Begin trailing spaces |
| 載入時發生錯誤 | General form load error — check .log file next to .frm |

## Validation

After creating or modifying a .frm file:

1. Check encoding: `file Source/myform.frm` → should show `ISO-8859 text, with CRLF`
2. Check Begin lines have trailing space: `grep "Begin " Source/myform.frm | cat -A`
3. Check Object lines have quotes: `grep "Object" Source/myform.frm`
4. Compile and check result
