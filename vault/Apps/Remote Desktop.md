---
aliases:
  - mstsc
  - RDP
  - Remote Desktop Connection
tags:
  - app
---
# Remote Desktop

The companion stays on top of a full-screen RDP session. Switch on **Clip**: text you copy inside the session is matched against this vault, so a copied `nginx` command brings up [[Dev/Linux Shell]].

The ini contains a commented rule `mstsc = Remote Desktop` that would always show this note for the RDP client.

## Shortcuts inside a session

| Keys | Action |
| --- | --- |
| Ctrl+Alt+End | Ctrl+Alt+Del on the remote machine |
| Ctrl+Alt+Break | Toggle full screen |
| Ctrl+Alt+Plus | Screenshot of the remote screen |
| Alt+Page Up | Switch programs (Alt+Tab) |

## Command line

```powershell
mstsc /v:server01 /f /admin
mstsc /v:server01 /w:1600 /h:900 /multimon
```
