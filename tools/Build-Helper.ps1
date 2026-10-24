#Requires -Version 7
<#
    .SYNOPSIS
    Builds dll\CheatSheetCompanionHelper.dll from the C# source embedded in CheatSheetCompanion.ps1 ($helperSource).

    .DESCRIPTION
    Run this once after editing the C# helpers (region "Native helpers"). The DLL gets an AssemblyVersion derived from
    the source hash (same rule as Get-HelperVersion in the script), so the app loads it only while it matches the
    source and otherwise compiles in memory. Nothing else is needed: the app itself never compiles to disk.

    .EXAMPLE
    pwsh -NoProfile -File .\tools\Build-Helper.ps1
#>
param([string]$ScriptPath = (Join-Path (Split-Path $PSScriptRoot -Parent) 'CheatSheetCompanion.ps1'))
$ErrorActionPreference = 'Stop'

function Get-HelperVersion {
    # identical to Get-HelperVersion in CheatSheetCompanion.ps1
    param([Parameter(Mandatory)][string]$Source)
    $sha = [System.Security.Cryptography.SHA256]::Create()
    $b = $sha.ComputeHash([System.Text.Encoding]::UTF8.GetBytes($Source.Replace("`r`n", "`n")))
    $sha.Dispose()
    return [version]::new([BitConverter]::ToUInt16($b, 0), [BitConverter]::ToUInt16($b, 2), [BitConverter]::ToUInt16($b, 4), [BitConverter]::ToUInt16($b, 6))
}

$tokens = $null; $errors = $null
$ast = [System.Management.Automation.Language.Parser]::ParseFile($ScriptPath, [ref]$tokens, [ref]$errors)
if ($errors.Count) { throw "Parse errors in $ScriptPath" }
$assign = $ast.Find({ param($n) $n -is [System.Management.Automation.Language.AssignmentStatementAst] -and $n.Left.Extent.Text -eq '$helperSource' }, $true)
if ($null -eq $assign) { throw 'Assignment $helperSource = @''...''@ not found' }
$source = [string]$assign.Right.Expression.Value
if (-not $source) { throw 'Helper source is empty' }

$version = Get-HelperVersion -Source $source
$stamped = [regex]::new('(?m)^namespace ').Replace($source, ('[assembly: System.Reflection.AssemblyVersion("' + $version + '")]' + "`n" + 'namespace '), 1)

$out = Join-Path (Split-Path $ScriptPath -Parent) 'dll\CheatSheetCompanionHelper.dll'
if (Test-Path -LiteralPath $out) { Remove-Item -LiteralPath $out -Force }   # fails while an instance of the app is running
Add-Type -TypeDefinition $stamped -OutputAssembly $out
$built = [System.Reflection.AssemblyName]::GetAssemblyName($out).Version
if ($built -ne $version) { throw "Version mismatch: built $built, expected $version" }
"Built {0} ({1} bytes), version {2}" -f $out, (Get-Item -LiteralPath $out).Length, $built
