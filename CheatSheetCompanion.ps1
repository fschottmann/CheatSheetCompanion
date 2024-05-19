#Requires -Version 7

<#PSScriptInfo

	.VERSION 2.0.0

	.GUID 1a0524f5-edba-4780-b562-6097e428b673

	.AUTHOR Frank Schottmann

	.COMPANYNAME

	.COPYRIGHT @2024 Frank Schottmann

	.TAGS 

	.LICENSEURI 

	.PROJECTURI 

	.ICONURI 

	.EXTERNALMODULEDEPENDENCIES 

	.REQUIREDSCRIPTS 

	.EXTERNALSCRIPTDEPENDENCIES 

	.RELEASENOTES
	Version 2.0.0
	- Code cleanup, including bug fixes.
	- Active check for new or changed markdown files (FileIndex).
	- Added picture support.
	- Active window grab starts automatically.
	- Added CosineSimilarity to find suitable files.
	- Displays which file was chosen for active window sync.
	
  Version 1.1.0
	- code cleanup
    - add WebView2
	Version 1.0.0
	- Initial Release
#>

<#
	.DESCRIPTION 
	Cheat Sheet Companion is a powerful PowerShell tool that provides context-sensitive cheat sheets to users. 
	No matter what program or website the user is in, the appropriate cheat sheets are always displayed, making 
	it easier for users to find the information they need quickly and efficiently.
	
#>


function Get-IniContent {
	[CmdletBinding()]
    param (
        [Parameter(Mandatory = $true,Position = 0)][Alias('Path')][string]$INIFilePath
    )
	
    begin
    {
        $INIContent = [Ordered]@{}
    
    }
    
    process
    {
        switch -regex ($(Get-Content $INIFilePath))
        {
            '^\[(.+)\]' # Section
            {
                $Section = $matches[1].Trim()
                $INIContent[$Section] = @{}
                $CommentCount = 0
            }
            "^(;.*)$" # Comment
            {
                $Value = $matches[1].Trim()
                if(-not (Get-Variable -Name 'CommentCount' -ErrorAction SilentlyContinue)) 
                {
                    $CommentCount = 1
                }
                else 
                {
                    $CommentCount = $CommentCount + 1
                }
                $Name = 'Comment' + $CommentCount
                if(Get-Variable -Name 'Section' -ErrorAction SilentlyContinue) 
                {
                    $INIContent[$Section][$Name] = $Value.Trim()
                }
            } 
            '(.+?)\s*=(.*)' # Key
            {
                $Name, $Value = $matches[1..2].Trim()
                $INIContent[$Section][$Name] = $Value.Trim()
            }
        }
    
    }
    
    end
    {
        return $INIContent
    
    }
}

function Get-FileHash {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory=$true)]
        [string]$FolderPath,

        [Parameter()]
        [string[]]$Filter = @('*.md', '*.png', '*.gif')
    )

    $files = Get-ChildItem -Path $FolderPath -File -Recurse -Include $Filter
    $sha256 = [System.Security.Cryptography.SHA256]::Create()
    $hashObjects = @()

    $files | ForEach-Object {
        $file = $_
        $fileStream = $file.OpenRead()
        $hashBytes = $sha256.ComputeHash($fileStream)
        $hashString = [System.BitConverter]::ToString($hashBytes).Replace("-", "")
        $fileStream.Close()

        $relativePath = $file.FullName.Replace($FolderPath + '\', '')

        $hashObject = [PSCustomObject]@{
            FileName     = $file.BaseName
            Hash         = $hashString
            RelativePath = $relativePath
        }

        $hashObjects += $hashObject
    }

    return $hashObjects
}

function Convert-ToWindowsFilename {
    param (
        [string]$filename
    )

    $invalidChars = [System.IO.Path]::GetInvalidFileNameChars()

    $sanitizedFilename = $filename
    foreach ($char in $invalidChars) {
        $sanitizedFilename = $sanitizedFilename -replace [Regex]::Escape($char), '_'
    }

    $sanitizedFilename = $sanitizedFilename.TrimEnd('. ')

    return $sanitizedFilename
}

function Convert-MarkdownToHtml {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory=$true)]
        [System.IO.FileInfo]$MarkdownFile,

        [Parameter(Mandatory=$true)]
        [string]$OutputPath
    )

    $conversionType = [Microsoft.PowerShell.MarkdownRender.MarkdownConversionType]::HTML
    $markdownOption = [Microsoft.PowerShell.MarkdownRender.PSMarkdownOptionInfo]::new()
    $markdownContent = [System.IO.File]::ReadAllText($MarkdownFile.FullName)
    $markdownConverter = [Microsoft.PowerShell.MarkdownRender.MarkdownConverter]
    $markdownResult = $markdownConverter::Convert($markdownContent, $conversionType, $markdownOption)
    $outputFilePath = Join-Path -Path $OutputPath -ChildPath ($MarkdownFile.BaseName + '.html')
    $markdownResult.Html | Out-File $outputFilePath -Force
}


function Get-NGrams {
    param (
        [string] $inputString,
        [int] $n
    )

    $inputString = -join ($inputString.ToLower() -replace '[^a-z0-9]', '')

    $length = $inputString.Length
    $ngrams = @()

    for ($i = 0; $i -le $length - $n; $i++) {
        $ngrams += $inputString.Substring($i, $n)
    }

    return $ngrams
}

function Get-CosineSimilarity {
    param (
        [string] $s,
        [string] $t,
        [int] $n
    )

    $sNGrams = Get-NGrams -inputString $s -n $n
    $tNGrams = Get-NGrams -inputString $t -n $n

    # Create frequency dictionaries
    $sFreq = @{}
    $tFreq = @{}

    foreach ($gram in $sNGrams) {
        if ($sFreq.ContainsKey($gram)) {
            $sFreq[$gram]++
        } else {
            $sFreq[$gram] = 1
        }
    }

    foreach ($gram in $tNGrams) {
        if ($tFreq.ContainsKey($gram)) {
            $tFreq[$gram]++
        } else {
            $tFreq[$gram] = 1
        }
    }

    $intersection = 0

    foreach ($gram in $sFreq.Keys) {
        if ($tFreq.ContainsKey($gram)) {
            $intersection += [math]::Min($sFreq[$gram], $tFreq[$gram])
        }
    }

    $similarity = 2 * $intersection / ($sNGrams.Count + $tNGrams.Count)

    return $similarity
}

function Get-MostSimilarFileName {
    param (
        [PSCustomObject[]] $FileIndex,
        [string] $targetString,
        [int] $n,
				[float] $threshold = 0.1  # Default similarity threshold
    )

    $maxSimilarity = -1
    $mostSimilarFile = $null

    $FileIndex | ForEach-Object {
        $file = $_
        $similarity = Get-CosineSimilarity -s $file.FileName -t $targetString -n $n

        if ($similarity -gt $maxSimilarity -and $similarity -ge $threshold) {
            $maxSimilarity = $similarity
            $mostSimilarFile = $file.FileName
        }
    }

    return $mostSimilarFile
}

function Convert-HTMLContent {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory=$false)]
        [string]$TitleContent = '',

        [Parameter(Mandatory=$false)]
        [string]$ProcessContent = '',
				
				[switch]$noTitle
    )

    begin {
    }

    process {
        # Process image syntax
        $imageSyntaxRegex = '!\[\[(.+?)\]\]'
        $replacementTemplate = '<img src="http://CSCImages/{0}">'
				
				if($TitleContent.count -gt 0) {
					[regex]::Matches($TitleContent, $imageSyntaxRegex) | ForEach-Object {
							$match = $_.Groups[1].Value
							$indexImage = $global:CheatSheet.FileIndex | Where-Object { $_.FileName -eq $match.split('.')[0] }
							if ($null -ne $indexImage) {
									$replacement = $replacementTemplate -f $indexImage.RelativePath.Replace('\', '/')
									$TitleContent = $TitleContent -replace [regex]::Escape($_.Value), $replacement
							}
					}
				}
				
        if($ProcessContent.count -gt 0) {
            [regex]::Matches($ProcessContent, $imageSyntaxRegex) | ForEach-Object {
                $match = $_.Groups[1].Value
                $indexImage = $global:CheatSheet.FileIndex | Where-Object { $_.FileName -eq $match.split('.')[0] }
                if ($null -ne $indexImage) {
                    $replacement = $replacementTemplate -f $indexImage.RelativePath.Replace('\', '/')
                    $ProcessContent = $ProcessContent -replace [regex]::Escape($_.Value), $replacement
                }
            }
        }

        # Process note links
        $noteLinkRegex = '\[\[([^#\[\]|]+)(?:#([^|\[\]]+))?(?:\|[^|\[\]]+)?\]\]'
				
        if($TitleContent.count -gt 0) {
					[regex]::Matches($TitleContent, $noteLinkRegex) | ForEach-Object {
							$match = $_
							$linkTarget = [System.Net.WebUtility]::HtmlEncode($match.Groups[1].Value)
							$linkHash = if ($match.Groups[2].Value) { "%23" + [System.Net.WebUtility]::HtmlEncode($match.Groups[2].Value) } else { "" }
							$replacement = '<a href="' + ($global:CheatSheet.General.INIConfig.Configuration.SourceLink -f $linkTarget) + $linkHash + '">' + $match.Groups[1].Value + $(if($match.groups[2].Value.Length -gt 0){' > ' + $match.groups[2].Value}) + '</a>'
							$TitleContent = $TitleContent -replace [regex]::Escape($match.Value), $replacement
					}
				}
				
        if($ProcessContent.count -gt 0) {
            [regex]::Matches($ProcessContent, $noteLinkRegex) | ForEach-Object {
                $match = $_
                $linkTarget = [System.Net.WebUtility]::HtmlEncode($match.Groups[1].Value)
                $linkHash = if ($match.Groups[2].Value) { "%23" + [System.Net.WebUtility]::HtmlEncode($match.Groups[2].Value) } else { "" }
                $replacement = '<a href="' + ($global:CheatSheet.General.INIConfig.Configuration.SourceLink -f $linkTarget) + $linkHash + '">' + $match.Groups[1].Value + $(if($match.groups[2].Value.Length -gt 0){' > ' + $match.groups[2].Value}) + '</a>'
                $ProcessContent = $ProcessContent -replace [regex]::Escape($match.Value), $replacement
            }
        }

        # Generate HTML content
        $htmlContent = @'

<html>
    <link rel="stylesheet" href="./styles/default.min.css">
    <script src="highlight.min.js"></script>
    <script>hljs.highlightAll();</script>
'@ + $(if (Test-Path "$($global:CheatSheet.General.CurrentPath)\styles\userstyle.css") { 
        '<style>' + (Get-Content -Path "$($global:CheatSheet.General.CurrentPath)\styles\userstyle.css" -Raw) + '</style>'
    }) +
@'    
    <body class="markdown-body">
'@ + $( if($TitleContent.Length -gt 0) {
@'		
        <div class="file-section">
'@ + $(if (-not ($noTitle)) {
        '<div class="file-header">found for title: ' +
				 $global:CheatSheet.LastActiveProcess.Files.Title +
				 '</div>
         <div class="splitter"></div>'
    }) +
@'
            <div class="file-content">
'@ + $TitleContent +
@'                    
            </div>
        </div>
'@  } ) +
@'				
'@ + $( if($ProcessContent.Length -gt 0) {
@'
        <div class="file-section">
            <div class="file-header">found for process: 
'@  + $global:CheatSheet.LastActiveProcess.Files.Process +
@'
            </div>
            <div class="splitter"></div>
            <div class="file-content">
'@  + $ProcessContent + 
@'
            </div>
        </div>
'@  } ) +
@'
   </body>
</html>
'@

        # Output the HTML content
        return $htmlContent
    }
}


Add-Type -AssemblyName `
    PresentationCore,`
    PresentationFrameWork,`
    WindowsBase,`
    System.Windows.Forms,`
    System.Security,`
    System.Drawing

Add-Type -TypeDefinition @'
    using System;
    using System.IO;
    using System.Diagnostics;
    using System.Runtime.InteropServices;
    using System.Text;
    using System.Windows.Forms;

    namespace CheatSheetCompanionHelper {

        public class WindowHelper {
            [DllImport("user32.dll", CharSet = CharSet.Auto, SetLastError = true)]
                public static extern int GetWindowText(IntPtr hwnd,StringBuilder lpString, int cch);

            [DllImport("user32.dll", SetLastError=true, CharSet=CharSet.Auto)]
                public static extern IntPtr GetForegroundWindow();

            [DllImport("user32.dll", SetLastError=true, CharSet=CharSet.Auto)]
                public static extern Int32 GetWindowThreadProcessId(IntPtr hWnd,out Int32 lpdwProcessId);

            [DllImport("user32.dll", SetLastError=true, CharSet=CharSet.Auto)]
                public static extern Int32 GetWindowTextLength(IntPtr hWnd);

            [DllImport("kernel32.dll", SetLastError = true, CharSet=CharSet.Auto)]
                public static extern bool SetDllDirectory(string lpPathName);
        }  

        public static class KeyLogger {
            private const int WH_KEYBOARD_LL = 13;
            private const int WM_KEYDOWN = 0x0100;
    
            private static HookProc hookProc = HookCallback;
            private static IntPtr hookId = IntPtr.Zero;
            private static int keyCode = 0;
    
            [DllImport("user32.dll", SetLastError = true)]
            private static extern IntPtr CallNextHookEx(IntPtr hhk, int nCode, IntPtr wParam, IntPtr lParam);
            
            [DllImport("user32.dll", SetLastError = true)]
            private static extern IntPtr SetWindowsHookEx(int idHook, HookProc lpfn, IntPtr hMod, uint dwThreadId);
    
            [DllImport("user32.dll", SetLastError = true)]
            [return: MarshalAs(UnmanagedType.Bool)]
            private static extern bool UnhookWindowsHookEx(IntPtr hhk);
    
            [DllImport("kernel32.dll", CharSet = CharSet.Auto, SetLastError = true)]
            private static extern IntPtr GetModuleHandle(string lpModuleName);
            
            [DllImport("user32.dll", CharSet = CharSet.Auto, ExactSpelling = true)]
            private static extern short GetKeyState(int keyCode);
            
            
            public static bool WaitForKey(int skey, int sModifier1 = 0, int sModifier2 = 0) {
                hookId = SetHook(hookProc);
                Application.Run();
                UnhookWindowsHookEx(hookId);
                
                if (sModifier2 == 0) {
                    if (sModifier1 == 0) {
                        if ((GetKeyState(skey) & 0x8000) != 0) {
                            return true;
                        }
                        else {
                            return false;
                        }
                    }
                    else if ((GetKeyState(sModifier1) & 0x8000) != 0 && (GetKeyState(skey) & 0x8000) != 0) {
                        return true;
                    }
                    else {
                        return false;
                    }
                }
                else if ((GetKeyState(sModifier2) & 0x8000) != 0 && (GetKeyState(sModifier1) & 0x8000) != 0 && (GetKeyState(skey) & 0x8000) != 0) {
                    return true;
                }
                else {
                    return false;
                }
            }
    
            private static IntPtr SetHook(HookProc hookProc) {
            IntPtr moduleHandle = GetModuleHandle(System.Environment.ProcessPath);
            return SetWindowsHookEx(WH_KEYBOARD_LL, hookProc, moduleHandle, 0);
            }
            
            private delegate IntPtr HookProc(int nCode, IntPtr wParam, IntPtr lParam);
            
            private static IntPtr HookCallback(int nCode, IntPtr wParam, IntPtr lParam) {
                if (nCode >= 0 && wParam == (IntPtr)WM_KEYDOWN) {
                    keyCode = Marshal.ReadInt32(lParam);
                    Application.Exit();
                }
                return CallNextHookEx(hookId, nCode, wParam, lParam);
            }
        }
    }
'@ -ReferencedAssemblies System.Windows.Forms


$global:CheatSheet = [pscustomobject]@{
    'General' = [pscustomobject]@{
        'WindowHeight' = 0
        'WindowType' = $null
        'CurrentPath' = $(if($PSScriptRoot.Length -lt 1){$pwd.Path}else{$PSScriptRoot})
        'INIConfig' = $(Get-IniContent -INIFilePath $($(if($PSScriptRoot.Length -lt 1){$pwd.Path}else{$PSScriptRoot}) + '\ini\config.ini'))
    }
    'Threads' = [pscustomobject]@{
        'ActiveWindow' = [pscustomobject]@{
            'PowerShell' = $null
            'Runspace' = $null
            'Handle' = $null
        }
        'GlobalShortcut' = [pscustomobject]@{
            'PowerShell' = $null
            'Runspace' = $null
            'Handle' = $null
        }
        'MarkDownConvert' = [pscustomobject]@{
            'PowerShell' = $null
            'Runspace' = $null
            'Handle' = $null
        }
    }
    'FileIndex' = $null
    'LastActiveProcess' = [pscustomobject]@{
        'Title' = $null
        'Process' = $null
        'Files' = [pscustomobject]@{
						'Title' = $null
						'Process' = $null
				}
    }
    'AutoSwitch' = [pscustomobject]@{
        'Active' = $true
        'SynchronizeWindow' = [hashtable]::Synchronized(@{})
        'LastProcess' = [pscustomobject]@{
            'Title' = ''
            'Process' = ''
        }
    }
}

if (($null -eq $global:CheatSheet.General.INIConfig.Configuration.MarkDownFolder) -or ($global:CheatSheet.General.INIConfig.Configuration.MarkDownFolder.Length -lt 1)){
    $global:CheatSheet.General.INIConfig += [PSCustomObject]@{
        'Configuration' = @([PSCustomObject]@{
            'MarkDownFolder' = $($global:CheatSheet.General.CurrentPath + '\md')
            'SourceLink' = ''
        })
    }
    $global:CheatSheet.General.INIConfig += [PSCustomObject]@{
        'RegEx' = @([PSCustomObject]@{
            'Prozess' = ''
            'NoteLink' = ''
        })
    }
}

<#
    Cleanup old converted markdown files
#>
if(-not $(Test-Path -Path $($global:CheatSheet.General.CurrentPath + '\Output'))) {
    New-Item -Name 'Output' -ItemType Directory -Path $global:CheatSheet.General.CurrentPath -Force
}
else {
    Get-ChildItem -Path $($global:CheatSheet.General.CurrentPath + '\Output\') | Remove-Item -Force -Recurse
}

<#
    Convert markdown files to HTML and create the file index
#>

Get-ChildItem -Path $global:CheatSheet.General.INIConfig.Configuration.MarkDownFolder -Filter '*.md' -Recurse | ForEach-Object {
    Convert-MarkdownToHtml -MarkdownFile $_.FullName -OutputPath $($global:CheatSheet.General.CurrentPath + '\Output')
}

$global:CheatSheet.FileIndex = Get-FileHash -FolderPath $global:CheatSheet.General.INIConfig.Configuration.MarkDownFolder

<#
    Check the existing browser engines
#>

if(Test-Path $($env:ProgramFiles + '\PackageManagement\NuGet\Packages\Microsoft.Web.WebView*')) {
	[CheatSheetCompanionHelper.WindowHelper]::SetDllDirectory($($(Split-Path $(Get-Item $($env:ProgramFiles + '\PackageManagement\NuGet\Packages\Microsoft.Web.WebView*\runtimes\win-x64\native\WebView2Loader.dll')).FullName -Parent)))
	Add-Type -Path $(Get-Item $($env:ProgramFiles + '\PackageManagement\NuGet\Packages\Microsoft.Web.WebView*\lib\netcoreapp3.0\Microsoft.Web.WebView2.Wpf.dll')).FullName
	Add-Type -Path $(Get-Item $($env:ProgramFiles + '\PackageManagement\NuGet\Packages\Microsoft.Web.WebView*\lib\netcoreapp3.0\Microsoft.Web.WebView2.Core.dll')).FullName

    $global:CheatSheet.General.WindowType = 'WebView2'
}
else {
    $global:CheatSheet.General.WindowType = 'WebBrowser'
}

<#
    Draw the WPF GUI
#>

$Window_XAML = @'
<Window 
    Title="CheatSheetCompanion" SizeToContent="Manual" Topmost="True" 
    xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
    xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
    xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"
'@

if($global:CheatSheet.General.WindowType -eq 'WebView2') {
    $Window_XAML += '    xmlns:Wpf="clr-namespace:Microsoft.Web.WebView2.Wpf;assembly=Microsoft.Web.WebView2.Wpf"'
}

$Window_XAML += @'
    Width="600" Height="450" WindowStyle="ToolWindow" Opacity="0.85">
    <Grid Name="MainGrid" Margin="5">
        <Border Name="MinimizeBorder" BorderThickness="1" BorderBrush="Gray" Margin="-10,-10,-20,-10" VerticalAlignment="Top">
            <Button Name="MinimizeButton">
                <TextBlock Name="MinimizeText" FontFamily="Segoe UI Symbol" Text="&#x1f53a;&#x1f53a;&#x1f53a;" FontSize="10"/>
            </Button>
        </Border>
        <Grid Name="InnerGrid">
            <Grid.RowDefinitions>
                <RowDefinition Height="Auto"/>
                <RowDefinition Height="*"/>
            </Grid.RowDefinitions>
            <Grid.ColumnDefinitions>
                <ColumnDefinition Width="*"/>
                <ColumnDefinition Width="Auto"/>
            </Grid.ColumnDefinitions>
            <StackPanel Grid.Column="0" Grid.Row="0" Orientation="Horizontal" Margin="5,15,5,0">
                <ComboBox Name="MarkdownComboBox" IsEditable="True" VerticalAlignment="Top" HorizontalAlignment="Left" Width="138" ToolTip="Search and select a note to view.">
'@

$global:CheatSheet.FileIndex | Where-Object {$_.RelativePath -match '.md'} | ForEach-Object {
    $Window_XAML += '                <ComboBoxItem>' + $_.FileName + '</ComboBoxItem>'
}

$Window_XAML += @'            
                </ComboBox>
            </StackPanel>    
            <StackPanel Grid.Column="1" Grid.Row="0" Orientation="Horizontal" Margin="5,15,5,0">
                <Button Background="LightGray" Name="btnEditMarkdown" BorderBrush="Gray" BorderThickness="1" Height="22" Width="25" HorizontalAlignment="Right"  Margin="0,0,5,0" VerticalAlignment="Top" ToolTip="Edit the selected note.">
                    <StackPanel Orientation="Horizontal">
                        <TextBlock FontFamily="Segoe UI Symbol" Text="&#xE104;" Name="txtEditMarkdown" HorizontalAlignment="Right"/>
                    </StackPanel>
                </Button>
                <Button Background="LightGray" Name="btnAddMarkdown" BorderBrush="Gray" BorderThickness="1" Height="22" Width="25" HorizontalAlignment="Right"  Margin="0,0,5,0" VerticalAlignment="Top" ToolTip="Create a new markdown according to the activated window.">
                    <StackPanel Orientation="Horizontal">
                        <TextBlock FontFamily="Segoe UI Symbol" Text="&#xE109;" Name="txtAddMarkdown" HorizontalAlignment="Right"/>
                    </StackPanel>
                </Button>
                <Button Background="LightGray" Name="btnLastActiveWindow" BorderBrush="Gray" BorderThickness="1" Height="22" Width="25" HorizontalAlignment="Right"  Margin="0,0,0,0" VerticalAlignment="Top" >
                    <TextBlock.ToolTip>
                        <TextBlock>
                            Select the last window. <Bold>Right-click</Bold>: Activate the automatic notes switching according to the active window.
                        </TextBlock>  
                    </TextBlock.ToolTip> 					
                    <StackPanel Orientation="Horizontal">
                        <TextBlock FontFamily="Segoe UI Symbol" Text="&#xE149;" Name="txtLastActiveWindow" Foreground="#FFED1C24" HorizontalAlignment="Right"/>
                    </StackPanel>
                </Button>
            </StackPanel>    
'@            
if($global:CheatSheet.General.WindowType -eq 'WebView2') {
    $Window_XAML += "`n" + '            <Wpf:WebView2 Name="WebBrowser" Source="about:blank" Margin="5,5,5,25" Grid.Row="1" Grid.ColumnSpan="2" />' + "`n"
} else {
    $Window_XAML += "`n" + '            <WebBrowser Name="WebBrowser" Margin="5,5,5,25" HorizontalAlignment="Left" VerticalAlignment="Bottom" Grid.Row="1" Grid.ColumnSpan="2" />' + "`n"
}


$Window_XAML += @'            
        </Grid>
        <StatusBar Name="StatusBar" Margin="5,15,5,0" HorizontalAlignment="Stretch" VerticalAlignment="Bottom">
            <StatusBarItem>
                <TextBlock Name="txtStatus" Text="Ready"/>
            </StatusBarItem>
        </StatusBar>
    </Grid>
</Window>
'@


$strWrite = New-Object IO.StringWriter
$([xml]$Window_XAML).save($strWrite)

$global:CheatSheet.AutoSwitch.SynchronizeWindow = [Windows.Markup.XamlReader]::Parse("$strWrite")
$window = $global:CheatSheet.AutoSwitch.SynchronizeWindow

Remove-Variable -Name Window_XAML -Force

<#
    Create the active file sync thread
#>

$initialSessionState = [System.Management.Automation.Runspaces.InitialSessionState]::CreateDefault()

$initialSessionState.Variables.Add(
        (New-object System.Management.Automation.Runspaces.SessionStateVariableEntry -ArgumentList "CheatSheet",$CheatSheet,$Null))

$definition = Get-Content Function:Get-FileHash -ErrorAction Stop
$addMessageSessionStateFunction = New-Object System.Management.Automation.Runspaces.SessionStateFunctionEntry -ArgumentList 'Get-FileHash', $definition
$initialSessionState.Commands.Add($addMessageSessionStateFunction)

$definition = Get-Content Function:Convert-MarkdownToHtml -ErrorAction Stop
$addMessageSessionStateFunction = New-Object System.Management.Automation.Runspaces.SessionStateFunctionEntry -ArgumentList 'Convert-MarkdownToHtml', $definition
$initialSessionState.Commands.Add($addMessageSessionStateFunction)

$global:CheatSheet.Threads.MarkDownConvert.Runspace = [System.Management.Automation.Runspaces.RunspaceFactory]::CreateRunspace($Host, $initialSessionState)
$($global:CheatSheet.Threads.MarkDownConvert.Runspace).ApartmentState = "STA"
$($global:CheatSheet.Threads.MarkDownConvert.Runspace).ThreadOptions = "ReuseThread"
$($global:CheatSheet.Threads.MarkDownConvert.Runspace).Open()


$Code = {
    Start-Sleep -Seconds 30

    while($true){
        $MarkDownChanges = Get-FileHash -FolderPath $global:CheatSheet.General.INIConfig.Configuration.MarkDownFolder | 
            Where-Object {$global:CheatSheet.FileIndex.Hash -notcontains $_.Hash}

        $NewMarkDownFiles = Get-FileHash -FolderPath $global:CheatSheet.General.INIConfig.Configuration.MarkDownFolder | 
            Where-Object {$global:CheatSheet.FileIndex.RelativePath -notcontains $_.RelativePath}

        $MissingMarkDownFiles = $global:CheatSheet.FileIndex | 
            Where-Object {-not (Test-Path -Path ($global:CheatSheet.General.INIConfig.Configuration.MarkDownFolder + '\' + $_.RelativePath))}

        if (($MarkDownChanges.count -gt 0) -or ($NewMarkDownFiles.count -gt 0) -or ($MissingMarkDownFiles.count -gt 0)) {
            $MarkDownChanges | ForEach-Object {
                $MarkDownChange = $_
                if($MarkDownChange.RelativePath -match '.md') {
                    Convert-MarkdownToHtml -MarkdownFile $($global:CheatSheet.General.INIConfig.Configuration.MarkDownFolder + '\' + $MarkDownChange.RelativePath) -OutputPath $($global:CheatSheet.General.CurrentPath + '\Output')
                }
                $($global:CheatSheet.FileIndex | Where-Object {$_.RelativePath -eq $MarkDownChange.RelativePath}).Hash = $MarkDownChange.Hash
            }

            $NewMarkDownFiles | ForEach-Object {
                $NewMarkDownFile = $_
                if($NewMarkDownFile.RelativePath -match '.md') {
                    Convert-MarkdownToHtml -MarkdownFile $($global:CheatSheet.General.INIConfig.Configuration.MarkDownFolder + '\' + $NewMarkDownFile.RelativePath) -OutputPath $($global:CheatSheet.General.CurrentPath + '\Output')
                }
                $global:CheatSheet.FileIndex += $NewMarkDownFile
            }

            $MissingMarkDownFiles | ForEach-Object {
                $MissingMarkDownFile = $_
                $global:CheatSheet.FileIndex = $global:CheatSheet.FileIndex | Where-Object {$_.RelativePath -ne $MissingMarkDownFile.RelativePath}
            }
        }

        Start-Sleep -Seconds 5
    }
}

$global:CheatSheet.Threads.MarkDownConvert.PowerShell = [PowerShell]::Create()
$($global:CheatSheet.Threads.MarkDownConvert.PowerShell).AddScript($Code) | Out-Null

$($global:CheatSheet.Threads.MarkDownConvert.PowerShell).Runspace = $global:CheatSheet.Threads.MarkDownConvert.Runspace
$global:CheatSheet.Threads.MarkDownConvert.Handle = $($global:CheatSheet.Threads.MarkDownConvert.PowerShell).BeginInvoke()

<#
    Create the active windows check thread
#>

$initialSessionState = [System.Management.Automation.Runspaces.InitialSessionState]::CreateDefault()

$initialSessionState.Variables.Add(
	(New-object System.Management.Automation.Runspaces.SessionStateVariableEntry -ArgumentList "CheatSheet",$global:CheatSheet,$Null))

$initialSessionState.Variables.Add(
  (New-object System.Management.Automation.Runspaces.SessionStateVariableEntry -ArgumentList "Window",$Window,$Null))

$initialSessionState.Variables.Add(
	(New-object System.Management.Automation.Runspaces.SessionStateVariableEntry -ArgumentList "ShowResult",$($global:CheatSheet.AutoSwitch.SynchronizeWindow.FindName("ShowResult")),$Null))

$definition = Get-Content Function:Get-MostSimilarFileName -ErrorAction Stop
$addMessageSessionStateFunction = New-Object System.Management.Automation.Runspaces.SessionStateFunctionEntry -ArgumentList 'Get-MostSimilarFileName', $definition
$initialSessionState.Commands.Add($addMessageSessionStateFunction)   

$definition = Get-Content Function:Get-CosineSimilarity -ErrorAction Stop
$addMessageSessionStateFunction = New-Object System.Management.Automation.Runspaces.SessionStateFunctionEntry -ArgumentList 'Get-CosineSimilarity', $definition
$initialSessionState.Commands.Add($addMessageSessionStateFunction)    

$definition = Get-Content Function:Get-NGrams -ErrorAction Stop
$addMessageSessionStateFunction = New-Object System.Management.Automation.Runspaces.SessionStateFunctionEntry -ArgumentList 'Get-NGrams', $definition
$initialSessionState.Commands.Add($addMessageSessionStateFunction)        

$definition = Get-Content Function:Convert-HTMLContent -ErrorAction Stop
$addMessageSessionStateFunction = New-Object System.Management.Automation.Runspaces.SessionStateFunctionEntry -ArgumentList 'Convert-HTMLContent', $definition
$initialSessionState.Commands.Add($addMessageSessionStateFunction)        

$global:CheatSheet.Threads.ActiveWindow.Runspace = [System.Management.Automation.Runspaces.RunspaceFactory]::CreateRunspace($Host, $initialSessionState)
$($global:CheatSheet.Threads.ActiveWindow.Runspace).ApartmentState = "STA"
$($global:CheatSheet.Threads.ActiveWindow.Runspace).ThreadOptions = "ReuseThread"
$($global:CheatSheet.Threads.ActiveWindow.Runspace).Open()


$Code = {
    Start-Sleep -Seconds 5

    while($true){
        $hWnd = [CheatSheetCompanionHelper.WindowHelper]::GetForegroundWindow()
        $len = [CheatSheetCompanionHelper.WindowHelper]::GetWindowTextLength($hWnd)
        $sb = New-Object text.stringbuilder -ArgumentList ($len + 1)
        $rtnlen = [CheatSheetCompanionHelper.WindowHelper]::GetWindowText($hWnd,$sb,$sb.Capacity)
        
        $global:CheatSheet.LastActiveProcess.Title = $sb.tostring()
        $processId = [uint]0 
        [CheatSheetCompanionHelper.WindowHelper]::GetWindowThreadProcessId($hWnd,[ref]$processId)
        $global:CheatSheet.LastActiveProcess.Process = $null
        $global:CheatSheet.LastActiveProcess.Process = [System.Diagnostics.Process]::GetProcessById($processId)
		
        if((($global:CheatSheet.AutoSwitch.Active -eq $true) -and ($global:CheatSheet.LastActiveProcess.Title -ne $global:CheatSheet.AutoSwitch.LastProcess.Title)) -and ($global:CheatSheet.LastActiveProcess.Title -ne 'CheatSheetCompanion')) {
						$HTMLContent = ''
						$n = 3
            $global:CheatSheet.LastActiveProcess.Files.Title = Get-MostSimilarFileName -FileIndex $global:CheatSheet.FileIndex -targetString $global:CheatSheet.LastActiveProcess.Title -n $n
            $global:CheatSheet.LastActiveProcess.Files.Process = Get-MostSimilarFileName -FileIndex $global:CheatSheet.FileIndex -targetString $global:CheatSheet.LastActiveProcess.Process -n $n
						
						if(($global:CheatSheet.LastActiveProcess.Files.Title.Length -gt 0) -or ($global:CheatSheet.LastActiveProcess.Files.Process.Length -gt 0)) {
							if($global:CheatSheet.LastActiveProcess.Files.Title -ne $global:CheatSheet.LastActiveProcess.Files.Process) {
								$TitleContent = Get-Content -Path $($global:CheatSheet.General.CurrentPath + '\Output\' + $global:CheatSheet.LastActiveProcess.Files.Title + '.html') -Raw
								$ProcessContent = Get-Content -Path $($global:CheatSheet.General.CurrentPath + '\Output\' + $global:CheatSheet.LastActiveProcess.Files.Process + '.html') -Raw
								
								$HTMLContent = Convert-HTMLContent -TitleContent $TitleContent -ProcessContent $ProcessContent
							}
							else {
								$TitleContent = Get-Content -Path $($global:CheatSheet.General.CurrentPath + '\Output\' + $global:CheatSheet.LastActiveProcess.Files.Title + '.html') -Raw
								$HTMLContent = Convert-HTMLContent -TitleContent $TitleContent
							}
						}
						
						if($HTMLContent.Length -gt 0) {
							$global:CheatSheet.AutoSwitch.SynchronizeWindow.Dispatcher.invoke(	[action]{
								
									$window.FindName("WebBrowser").CoreWebView2.SetVirtualHostNameToFolderMapping("CSCImages", $global:CheatSheet.General.INIConfig.Configuration.MarkDownFolder, [Microsoft.Web.WebView2.Core.CoreWebView2HostResourceAccessKind]::Allow)
									$Window.FindName('WebBrowser').NavigateToString($HTMLContent)
							})
						}
						
            $global:CheatSheet.AutoSwitch.LastProcess.Title = $sb.tostring()
            $global:CheatSheet.AutoSwitch.LastProcess.Process = $null
            $global:CheatSheet.AutoSwitch.LastProcess.Process = [System.Diagnostics.Process]::GetProcessById($processId)
                
            $global:CheatSheet.AutoSwitch.SynchronizeWindow.Dispatcher.invoke(	[action]{
                $Window.FindName('txtStatus').Text = $('Process: ' + $global:CheatSheet.LastActiveProcess.Process.Name +  ' - Title: ' + $global:CheatSheet.LastActiveProcess.Title)
            })                
        }
        
        Start-Sleep -Seconds 1
    }
}

$global:CheatSheet.Threads.ActiveWindow.PowerShell = [PowerShell]::Create()
$($global:CheatSheet.Threads.ActiveWindow.PowerShell).AddScript($Code) | Out-Null

$($global:CheatSheet.Threads.ActiveWindow.PowerShell).Runspace = $global:CheatSheet.Threads.ActiveWindow.Runspace
$global:CheatSheet.Threads.ActiveWindow.Handle = $($global:CheatSheet.Threads.ActiveWindow.PowerShell).BeginInvoke()


$Window.FindName('MinimizeButton').`
    add_Click({
        if($Window.Height -lt 15) {
            $Window.Height = $global:CheatSheet.General.WindowHeight

            $Window.FindName('InnerGrid').Visibility = 'Visible'
            $Window.FindName('StatusBar').Visibility = 'Visible'
            $Window.WindowStyle = [System.Windows.WindowStyle]::ToolWindow
            $Window.ResizeMode = [System.Windows.ResizeMode]::CanResize

            $Window.FindName('MinimizeText').Text = [System.Net.WebUtility]::HtmlDecode('&#x1f53a;&#x1f53a;&#x1f53a;') 
        }
        else {
            $global:CheatSheet.General.WindowHeight = $Window.Height

            $Window.FindName('InnerGrid').Visibility = 'Hidden'
            $Window.FindName('StatusBar').Visibility = 'Hidden'
            $Window.FindName('MinimizeText').Text = [System.Net.WebUtility]::HtmlDecode('&#x1f53b;&#x1f53b;&#x1f53b;')            
            $Window.WindowStyle = [System.Windows.WindowStyle]::None
            $Window.ResizeMode = [System.Windows.ResizeMode]::NoResize
            $Window.Height = '13'
        }
    })

$Window.FindName('MinimizeBorder').`
    Add_MouseDown({
        if([System.Windows.Forms.Control]::MouseButtons -eq [System.Windows.Forms.MouseButtons]::Right) {
            $Window.Close()
        }
    })

$Window.FindName('MarkdownComboBox').`
    Add_DropDownClosed({
        $selectedItem = $null
        $selectedItem = $Window.FindName('MarkdownComboBox').items | Where-Object {$_.IsSelected -eq $true}

        if($null -ne $selectedItem) {         
            
            $TitleContent = Get-Content -Path $($global:CheatSheet.General.CurrentPath + '\Output\' + $selectedItem.content + '.html') -Raw
						
						$global:CheatSheet.LastActiveProcess.Files.Title = $selectedItem.content
            $global:CheatSheet.LastActiveProcess.Files.Process = ''
						
						<#
							add a always content switch
						#>
						
						$global:CheatSheet.AutoSwitch.LastProcess.Title = $null
						$global:CheatSheet.AutoSwitch.LastProcess.Process = $null
						
            $window.FindName("WebBrowser").CoreWebView2.SetVirtualHostNameToFolderMapping("CSCImages", $global:CheatSheet.General.INIConfig.Configuration.MarkDownFolder, [Microsoft.Web.WebView2.Core.CoreWebView2HostResourceAccessKind]::Allow)
            $Window.FindName('WebBrowser').NavigateToString($(Convert-HTMLContent -TitleContent $TitleContent -noTitle))
        }
    }) 

$Window.FindName('MarkdownComboBox').`
    Add_SelectionChanged({
				$selectedItem = $null
				$selectedItem = $Window.FindName('MarkdownComboBox').items | Where-Object {$_.IsSelected -eq $true}

				if($null -ne $selectedItem) {         
						
						$TitleContent = Get-Content -Path $($global:CheatSheet.General.CurrentPath + '\Output\' + $selectedItem.content + '.html') -Raw
												
												$global:CheatSheet.LastActiveProcess.Files.Title = $selectedItem.content
						$global:CheatSheet.LastActiveProcess.Files.Process = ''
												
												<#
														add a always content switch
												#>
												
												$global:CheatSheet.AutoSwitch.LastProcess.Title = $null
												$global:CheatSheet.AutoSwitch.LastProcess.Process = $null
												
						$window.FindName("WebBrowser").CoreWebView2.SetVirtualHostNameToFolderMapping("CSCImages", $global:CheatSheet.General.INIConfig.Configuration.MarkDownFolder, [Microsoft.Web.WebView2.Core.CoreWebView2HostResourceAccessKind]::Allow)
						$Window.FindName('WebBrowser').NavigateToString($(Convert-HTMLContent -TitleContent $TitleContent -noTitle))
				}
    }) 
		
$Window.FindName('txtEditMarkdown').`
    Add_MouseDown({
        if($global:CheatSheet.LastActiveProcess.Files.Title.Length -gt 0){

            if($global:CheatSheet.General.INIConfig.Configuration.SourceLink -gt 0) {
                Start-Process $($global:CheatSheet.General.INIConfig.Configuration.SourceLink -f $global:CheatSheet.LastActiveProcess.Files.Title) | Out-Null
            }
            else {
                Start-Process -FilePath $($global:CheatSheet.General.INIConfig.Configuration.MarkDownFolder + '\' + $global:CheatSheet.LastActiveProcess.Files.Title)
            }
        }
    })		

$Window.FindName('txtAddMarkdown').`
    Add_MouseDown({    
        if(-not(Test-Path $($global:CheatSheet.General.INIConfig.Configuration.MarkDownFolder + '\' + $global:CheatSheet.AutoSwitch.LastProcess.Title + '.md'))) {
            New-Item -Path $global:CheatSheet.General.INIConfig.Configuration.MarkDownFolder -Name $(Convert-ToWindowsFilename -filename $($global:CheatSheet.AutoSwitch.LastProcess.Title + '.md'))

            if($global:CheatSheet.General.INIConfig.Configuration.SourceLink -gt 0) {
                Start-Process $($global:CheatSheet.General.INIConfig.Configuration.SourceLink -f $(Convert-ToWindowsFilename -filename $($global:CheatSheet.AutoSwitch.LastProcess.Title))) | Out-Null
            }
            else {
                Start-Process -FilePath $($global:CheatSheet.General.INIConfig.Configuration.MarkDownFolder + '\' + $(Convert-ToWindowsFilename -filename $($global:CheatSheet.AutoSwitch.LastProcess.Title + '.md')))
            }
        }
    })		
		
$Window.FindName('txtLastActiveWindow').`
    Add_MouseDown({
            if($global:CheatSheet.AutoSwitch.Active -eq $true) {
                $global:CheatSheet.AutoSwitch.Active = $false
                $Window.FindName('txtLastActiveWindow').Foreground = '#FF000000'
                $Window.FindName('txtStatus').Text = 'Auto Sync deactivated.'                
            }
            else {
                $global:CheatSheet.AutoSwitch.Active = $true
                $Window.FindName('txtLastActiveWindow').Foreground = '#ED1C24'
								$Window.FindName('txtStatus').Text = 'Auto Sync activated.'
            }

            $global:CheatSheet.Threads.ActiveWindow.PowerShell.Stop()
            $global:CheatSheet.Threads.ActiveWindow.PowerShell.Streams.ClearStreams()
					
            $global:CheatSheet.Threads.ActiveWindow.Runspace.InitialSessionState.Variables.Remove("CheatSheet",$null)

            $global:CheatSheet.Threads.ActiveWindow.Runspace.InitialSessionState.Variables.Remove("Window",$null)
            
            $global:CheatSheet.Threads.ActiveWindow.Runspace.InitialSessionState.Variables.Add(
                (New-object System.Management.Automation.Runspaces.SessionStateVariableEntry -ArgumentList "CheatSheet",$global:CheatSheet,$Null)
            )

            $global:CheatSheet.Threads.ActiveWindow.Runspace.InitialSessionState.Variables.Add(
						(New-object System.Management.Automation.Runspaces.SessionStateVariableEntry -ArgumentList "ShowResult",$($global:CheatSheet.AutoSwitch.SynchronizeWindow.FindName("ShowResult")),$Null))

            $global:CheatSheet.Threads.ActiveWindow.Runspace.InitialSessionState.Variables.Add(
                (New-object System.Management.Automation.Runspaces.SessionStateVariableEntry -ArgumentList "Window",$Window,$Null)
            )

            $global:CheatSheet.Threads.ActiveWindow.PowerShell.Runspace.ResetRunspaceState()
            $global:CheatSheet.Threads.ActiveWindow.Handle = $($global:CheatSheet.Threads.ActiveWindow.PowerShell).BeginInvoke()						
   })        

if($global:CheatSheet.General.WindowType -ne 'WebView2') {   
    $Window.FindName('WebBrowser').Add_Navigating({
        param($sender, $e)

        $uri = $e.Uri
        if ($null -ne $uri) {    
            Start-Process $uri.AbsoluteUri
            $e.Cancel = $True
        }
    })    
} else {
    $handler = [System.EventHandler[Microsoft.Web.WebView2.Core.CoreWebView2NavigationStartingEventArgs]]{
        param($sender, $e)
        if ($($e.uri -notlike 'data:text/html*') -and ($e.uri -ne 'about:blank')) { 
            Start-Process $e.Uri
            $sender.stop()
        }
    }
    
    $Window.FindName('WebBrowser').add_NavigationStarting($handler)
}

$window.add_Loaded({

    $initialSessionState = [System.Management.Automation.Runspaces.InitialSessionState]::CreateDefault()

    $initialSessionState.Variables.Add(
		(New-object System.Management.Automation.Runspaces.SessionStateVariableEntry -ArgumentList "CheatSheet",$global:CheatSheet,$Null))

    $initialSessionState.Variables.Add(
        (New-object System.Management.Automation.Runspaces.SessionStateVariableEntry -ArgumentList "Window",$Window,$Null))

    $global:CheatSheet.Threads.GlobalShortcut.Runspace = [System.Management.Automation.Runspaces.RunspaceFactory]::CreateRunspace($Host, $initialSessionState)
    $($global:CheatSheet.Threads.GlobalShortcut.Runspace).ApartmentState = "STA"
    $($global:CheatSheet.Threads.GlobalShortcut.Runspace).ThreadOptions = "ReuseThread"
    $($global:CheatSheet.Threads.GlobalShortcut.Runspace).Open()

    $Code = {
        while ($true) {
            if($(Invoke-Command $([System.Management.Automation.ScriptBlock]::Create('[CheatSheetCompanionHelper.KeyLogger]::WaitForKey(' + $($($($global:CheatSheet.General.INIConfig.GlobalShortcut.DisplayCompanion -split ',') | ForEach-Object {[System.Management.Automation.ScriptBlock]::Create($('[System.Windows.Forms.Keys]::' + $_))}) -join ',') +')')) -NoNewScope)) {
                $global:CheatSheet.AutoSwitch.SynchronizeWindow.Dispatcher.invoke(	[action]{
                    if($Window.Height -lt 15) {
                        $Window.Height = $global:CheatSheet.General.WindowHeight
                        
            
                        $Window.FindName('InnerGrid').Visibility = 'Visible'
                        $Window.FindName('StatusBar').Visibility = 'Visible'
                        $Window.WindowStyle = [System.Windows.WindowStyle]::ToolWindow
                        $Window.ResizeMode = [System.Windows.ResizeMode]::CanResize
            
                        $Window.FindName('MinimizeText').Text = [System.Net.WebUtility]::HtmlDecode('&#x1f53a;&#x1f53a;&#x1f53a;') 
                    }
                    else {
                        $global:CheatSheet.General.WindowHeight = $Window.Height
            
                        $Window.FindName('InnerGrid').Visibility = 'Hidden'
                        $Window.FindName('StatusBar').Visibility = 'Hidden'
                        $Window.FindName('MinimizeText').Text = [System.Net.WebUtility]::HtmlDecode('&#x1f53b;&#x1f53b;&#x1f53b;')            
                        $Window.WindowStyle = [System.Windows.WindowStyle]::None
                        $Window.ResizeMode = [System.Windows.ResizeMode]::NoResize
                        $Window.Height = '13'
                    }
                })
            }
            elseif($(Invoke-Command $([System.Management.Automation.ScriptBlock]::Create('[CheatSheetCompanionHelper.KeyLogger]::WaitForKey(' + $($($($global:CheatSheet.General.INIConfig.GlobalShortcut.AutoSwitch -split ',') | ForEach-Object {[System.Management.Automation.ScriptBlock]::Create($('[System.Windows.Forms.Keys]::' + $_))}) -join ',') +')')) -NoNewScope)) {
                $global:CheatSheet.AutoSwitch.SynchronizeWindow.Dispatcher.invoke(	[action]{
                    if($global:CheatSheet.AutoSwitch.Active -eq $true) {
                        $global:CheatSheet.AutoSwitch.Active = $false
                        $Window.FindName('txtLastActiveWindow').Foreground = '#FF000000'
                        
                        $Window.FindName('txtStatus').Text = 'Auto Sync deactivated.'
                        
                    }
                    else {
                        $global:CheatSheet.AutoSwitch.Active = $true
                        $Window.FindName('txtLastActiveWindow').Foreground = '#ED1C24'
                    }

                    $global:CheatSheet.Threads.ActiveWindow.PowerShell.Stop()
                    $global:CheatSheet.Threads.ActiveWindow.PowerShell.Streams.ClearStreams()
                            
                    $global:CheatSheet.Threads.ActiveWindow.Runspace.InitialSessionState.Variables.Remove("CheatSheet",$null)

                    $global:CheatSheet.Threads.ActiveWindow.Runspace.InitialSessionState.Variables.Remove("Window",$null)
                    
                    $global:CheatSheet.Threads.ActiveWindow.Runspace.InitialSessionState.Variables.Add(
                        (New-object System.Management.Automation.Runspaces.SessionStateVariableEntry -ArgumentList "CheatSheet",$global:CheatSheet,$Null)
                    )

                    $global:CheatSheet.Threads.ActiveWindow.Runspace.InitialSessionState.Variables.Add(
                                (New-object System.Management.Automation.Runspaces.SessionStateVariableEntry -ArgumentList "ShowResult",$($global:CheatSheet.AutoSwitch.SynchronizeWindow.FindName("ShowResult")),$Null))

                    $global:CheatSheet.Threads.ActiveWindow.Runspace.InitialSessionState.Variables.Add(
                        (New-object System.Management.Automation.Runspaces.SessionStateVariableEntry -ArgumentList "Window",$Window,$Null)
                    )

                    $global:CheatSheet.Threads.ActiveWindow.PowerShell.Runspace.ResetRunspaceState()
                    $global:CheatSheet.Threads.ActiveWindow.Handle = $($global:CheatSheet.Threads.ActiveWindow.PowerShell).BeginInvoke()
                
                })	
            }
        }
    }

    $global:CheatSheet.Threads.GlobalShortcut.PowerShell = [PowerShell]::Create()
    $($global:CheatSheet.Threads.GlobalShortcut.PowerShell).AddScript($Code) | Out-Null

    $($global:CheatSheet.Threads.GlobalShortcut.PowerShell).Runspace = $global:CheatSheet.Threads.GlobalShortcut.Runspace
    $global:CheatSheet.Threads.GlobalShortcut.Handle = $($global:CheatSheet.Threads.GlobalShortcut.PowerShell).BeginInvoke()

})


$window.Add_SourceInitialized({
    $window.Add_LocationChanged({
        $sensitivity = 35
        if ($window.Top -lt $sensitivity) {
            $window.Top = 0
        }
    })
})

if($global:CheatSheet.General.WindowType -eq 'WebView2') {
    $window.FindName("WebBrowser").CreationProperties = New-Object Microsoft.Web.WebView2.Wpf.CoreWebView2CreationProperties
    $window.FindName("WebBrowser").CreationProperties.UserDataFolder = "$env:LocalAppData\CheatSheetCompanion"
}    

$window.ShowDialog()

$CheatSheet.Threads.ActiveWindow.PowerShell.Stop()
$CheatSheet.Threads.ActiveWindow.Handle = $null
$CheatSheet.Threads.ActiveWindow.PowerShell = $null
$CheatSheet.Threads.ActiveWindow.Runspace.Close()
$CheatSheet.Threads.ActiveWindow.Runspace.Dispose()
$CheatSheet.Threads.ActiveWindow.Runspace = $null

$CheatSheet.Threads.GlobalShortcut.PowerShell.Stop()
$CheatSheet.Threads.GlobalShortcut.Handle = $null
$CheatSheet.Threads.GlobalShortcut.PowerShell = $null
$CheatSheet.Threads.GlobalShortcut.Runspace.Close()
$CheatSheet.Threads.GlobalShortcut.Runspace.Dispose()
$CheatSheet.Threads.GlobalShortcut.Runspace = $null


if($null -ne $CheatSheet.Threads.MarkDownConvert.PowerShell) {
    $CheatSheet.Threads.MarkDownConvert.PowerShell.Stop()
    $CheatSheet.Threads.MarkDownConvert.Handle = $null
    $CheatSheet.Threads.MarkDownConvert.PowerShell = $null
    $CheatSheet.Threads.MarkDownConvert.Runspace.Close()
    $CheatSheet.Threads.MarkDownConvert.Runspace.Dispose()
    $CheatSheet.Threads.MarkDownConvert.Runspace = $null
}

[GC]::Collect()
[GC]::WaitForPendingFinalizers()		