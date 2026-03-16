---
name: compile_vb6_rc
description: Compile a VB6 resource file (.rc → .RES) using the plugin's compile_rc.bat
disable-model-invocation: true
argument-hint: "[resource.rc]"
---

Compile a VB6 resource file (.rc → .RES).

## Determine which .rc to compile

- If the user provided an argument: use `$ARGUMENTS` as the .rc file path.
- If no argument was provided: find all `.rc` files in the project directory and choose the most appropriate one. If there is only one, use it. If there are multiple, pick the primary one based on file name and project context.

## Compile

Run the compile command:

```
cmd.exe //c "${CLAUDE_PLUGIN_ROOT}\scripts\compile_rc.bat \"<FULL_PATH_TO_RC_FILE>\""
```

## Handle result

- If compilation succeeded — report success.
- If compilation failed — report the error and diagnose the issue.
