---
aliases:
  - pwsh
  - Windows PowerShell
  - Windows Terminal
tags:
  - dev
  - shell
---
# PowerShell

Found through the window titles of Windows Terminal, PowerShell 7 and Windows PowerShell (aliases) and through the process names `pwsh` and `powershell`.

## Everyday

```powershell
Get-Command -Noun Service              # find cmdlets
Get-Help Get-Process -Examples
Get-Process | Sort-Object CPU -Descending | Select-Object -First 5
Get-ChildItem -Recurse -Filter *.log | Where-Object Length -gt 1MB
Get-Content .\app.log -Tail 50 -Wait
```

## Text and objects

```powershell
'a,b,c' -split ','
'Hello' -replace 'l+', 'L'
Select-String -Path *.ps1 -Pattern 'TODO' -SimpleMatch
$obj | ConvertTo-Json -Depth 4
Import-Csv .\users.csv -Delimiter ';' | Group-Object Department
```

## Remoting and system

```powershell
Enter-PSSession -ComputerName server01
Invoke-Command -ComputerName server01 -ScriptBlock { Get-Service w32time }
Get-WinEvent -LogName System -MaxEvents 20
Test-NetConnection server01 -Port 3389
Restart-Service Spooler -Force
```

## Operators

| Operator | Meaning |
| --- | --- |
| `-eq -ne -gt -lt` | Comparison (case insensitive) |
| `-like -match` | Wildcard / regex (see [[Regex]]) |
| `-contains -in` | Collection membership |
| `?.` `??` | Null conditional / null coalescing (PS 7) |

See also [[Git]] and [[Linux Shell]].
