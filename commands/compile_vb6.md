---
name: compile_vb6
description: Compile a VB6 project (.vbp) using the plugin's compile.bat
disable-model-invocation: true
argument-hint: "[project.vbp]"
---

Compile a VB6 project.

## Determine which .vbp to compile

- If the user provided an argument: use `$ARGUMENTS` as the .vbp file path.
- If no argument was provided: find all `.vbp` files in the project directory and choose the most appropriate one (e.g., the main project, not test/utility projects). If there is only one, use it. If there are multiple, pick the primary one based on file name and project context.

## Compile

Run the compile command:

```
cmd.exe //c "cd /d <PROJECT_DIR> && ${CLAUDE_PLUGIN_ROOT}\scripts\compile.bat <VBP_FILE>"
```

## Handle result

- If the output contains `[OK] Compile succeeded.` — report success.
- If the output contains `[FAIL]` — read the `compile.log` and any `.log` files next to failing `.frm` files, diagnose the error, and report the issue with the relevant file and line number.

Remember: VB6 compile errors in `.frm` files report line numbers relative to the code section (after `Attribute VB_Exposed`), not the absolute file line.
