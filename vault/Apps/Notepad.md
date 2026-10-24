---
aliases:
  - Notepad++
  - Editor
tags:
  - app
---
# Notepad

Found through the window title of Notepad or Notepad++ (alias). Open one of them and watch the tab appear.

## Notepad++ shortcuts

| Keys | Action |
| --- | --- |
| Ctrl+D | Duplicate line |
| Ctrl+Shift+Up/Down | Move line |
| Ctrl+F2 | Toggle bookmark, F2 jumps |
| Ctrl+H | Replace (regex mode in the dialog) |
| Alt+Shift+Arrows | Column selection |
| Ctrl+Alt+Shift+R | Trim trailing spaces and save (plugin) |

## Regex replace examples

- Remove empty lines: find `^\s*$\r?\n`, replace with nothing
- Add a prefix to every line: find `^`, replace with `- `
- See [[Dev/Regex]] for the syntax.
