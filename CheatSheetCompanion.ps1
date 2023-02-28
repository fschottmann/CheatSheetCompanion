#Requires -Version 7

<#PSScriptInfo

	.VERSION 1.0.0

	.GUID 1a0524f5-edba-4780-b562-6097e428b673

	.AUTHOR Frank Schottmann

	.COMPANYNAME

	.COPYRIGHT @2023 Frank Schottmann

	.TAGS 

	.LICENSEURI 

	.PROJECTURI 

	.ICONURI 

	.EXTERNALMODULEDEPENDENCIES 

	.REQUIREDSCRIPTS 

	.EXTERNALSCRIPTDEPENDENCIES 

	.RELEASENOTES
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
	
	# https://cef-builds.spotifycdn.com/index.html
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

$global:CheatSheat = [pscustomobject]@{
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
    'LastActiveProcess' = [pscustomobject]@{
        'Title' = $null
        'Process' = $null
    }
    'AutoSwitch' = [pscustomobject]@{
        'Active' = $false
        'SynchronizeWindow' = [hashtable]::Synchronized(@{})
        'LastProcess' = [pscustomobject]@{
            'Title' = $null
            'Process' = $null
        }
    }
}


if(-not $(Test-Path -Path $($global:CheatSheat.General.CurrentPath + '\Output'))) {
    New-Item -Name 'Output' -ItemType Directory -Path $global:CheatSheat.General.CurrentPath -Force
}

if(-not $(Test-Path -Path $($global:CheatSheat.General.CurrentPath + '\md'))) {
    New-Item -Name 'md' -ItemType Directory -Path $global:CheatSheat.General.CurrentPath -Force
}

Get-ChildItem -Path $($global:CheatSheat.General.CurrentPath + '\md') -Filter '*.md' -Recurse | ForEach-Object {
    $(ConvertFrom-Markdown -Path $_.FullName).Html | Out-File $($global:CheatSheat.General.CurrentPath + '\Output\' + $_.BaseName + '.html') -Force
}

if(Test-Path $($env:ProgramFiles + '\PackageManagement\NuGet\Packages\Microsoft.Web.WebView*')) {
	[CheatSheetCompanionHelper.WindowHelper]::SetDllDirectory($($(Split-Path $(Get-Item $($env:ProgramFiles + '\PackageManagement\NuGet\Packages\Microsoft.Web.WebView*\runtimes\win-x64\native\WebView2Loader.dll')).FullName -Parent)))
	Add-Type -Path $(Get-Item $($env:ProgramFiles + '\PackageManagement\NuGet\Packages\Microsoft.Web.WebView*\lib\netcoreapp3.0\Microsoft.Web.WebView2.Wpf.dll')).FullName
	Add-Type -Path $(Get-Item $($env:ProgramFiles + '\PackageManagement\NuGet\Packages\Microsoft.Web.WebView*\lib\netcoreapp3.0\Microsoft.Web.WebView2.Core.dll')).FullName

    $global:CheatSheat.General.WindowType = 'WebView2'
}
else {
    $global:CheatSheat.General.WindowType = 'WebBrowser'
}


$initialSessionState = [System.Management.Automation.Runspaces.InitialSessionState]::CreateDefault()

$initialSessionState.Variables.Add(
		(New-object System.Management.Automation.Runspaces.SessionStateVariableEntry -ArgumentList "CheatSheat",$CheatSheat,$Null))

$initialSessionState.Variables.Add(
        (New-object System.Management.Automation.Runspaces.SessionStateVariableEntry -ArgumentList "Window",$null,$Null))

$global:CheatSheat.Threads.ActiveWindow.Runspace = [System.Management.Automation.Runspaces.RunspaceFactory]::CreateRunspace($Host, $initialSessionState)
$($global:CheatSheat.Threads.ActiveWindow.Runspace).ApartmentState = "STA"
$($global:CheatSheat.Threads.ActiveWindow.Runspace).ThreadOptions = "ReuseThread"
$($global:CheatSheat.Threads.ActiveWindow.Runspace).Open()


$Code = {
    while($true){
        $hWnd = [CheatSheetCompanionHelper.WindowHelper]::GetForegroundWindow()
        $len = [CheatSheetCompanionHelper.WindowHelper]::GetWindowTextLength($hWnd)
        $sb = New-Object text.stringbuilder -ArgumentList ($len + 1)
        $rtnlen = [CheatSheetCompanionHelper.WindowHelper]::GetWindowText($hWnd,$sb,$sb.Capacity)
        
        $global:CheatSheat.LastActiveProcess.Title = $sb.tostring()
        $processId = [uint]0 
        [CheatSheetCompanionHelper.WindowHelper]::GetWindowThreadProcessId($hWnd,[ref]$processId)
        $global:CheatSheat.LastActiveProcess.Process = $null
        $global:CheatSheat.LastActiveProcess.Process = [System.Diagnostics.Process]::GetProcessById($processId)
		
        if((($global:CheatSheat.AutoSwitch.Active -eq $true) -and ($global:CheatSheat.LastActiveProcess.Title -ne $global:CheatSheat.AutoSwitch.LastProcess.Title)) -and ($global:CheatSheat.LastActiveProcess.Title -ne 'CheatSheetCompanion')) {
            $file = ''

            Get-ChildItem -Path $($global:CheatSheat.General.CurrentPath + '\Output') | Sort-Object LastWriteTime -Descending | ForEach-Object {
                $item = $_  
                
                if($($($item.BaseName -split '_') -split '-' | Where-Object { $global:CheatSheat.LastActiveProcess.Title -like $('*' + $($_.Trim() + '*') ) }).Length -gt 0) {
                    $file += Get-Content -Path $item.FullName -Raw
                }
                elseif($($($item.BaseName -split '_') -split '-' | Where-Object { $global:CheatSheat.LastActiveProcess.Process.Name -like $('*' + $($_.Trim() + '*') ) }).Length -gt 0) {
                    $file += Get-Content -Path $item.FullName -Raw
                }
            }

            if($file.length -gt 0) {
                $global:CheatSheat.AutoSwitch.SynchronizeWindow.Dispatcher.invoke(	[action]{
                    $Window.FindName('WebBrowser').NavigateToString($(
                        '<link rel="stylesheet" href="./styles/default.min.css">' + `
                        '<script src="highlight.min.js"></script>' + `
                        '<script>hljs.highlightAll();</script>' + `
                        $(if(Test-Path $($global:CheatSheat.General.CurrentPath + '\styles\userstyle.css')){ 
                            '<style>' + $(Get-Content -Path $($global:CheatSheat.General.CurrentPath + '\styles\userstyle.css') -Raw) + '</style>'
                        })  + `
                        '<body class="markdown-body">' + $file + '</body></html>'
                    ))
                })
            }
            Remove-Variable -Name 'file' -Force

            $global:CheatSheat.AutoSwitch.LastProcess.Title = $sb.tostring()
            $global:CheatSheat.AutoSwitch.LastProcess.Process = $null
            $global:CheatSheat.AutoSwitch.LastProcess.Process = [System.Diagnostics.Process]::GetProcessById($processId)

                
            $global:CheatSheat.AutoSwitch.SynchronizeWindow.Dispatcher.invoke(	[action]{
                $Window.FindName('txtStatus').Text = $('Process: ' + $global:CheatSheat.LastActiveProcess.Process.Name +  ' - Title: ' + $global:CheatSheat.LastActiveProcess.Title)
            })                
        }
        
        Start-Sleep -Seconds 1
    }
}

$global:CheatSheat.Threads.ActiveWindow.PowerShell = [PowerShell]::Create()
$($global:CheatSheat.Threads.ActiveWindow.PowerShell).AddScript($Code) | Out-Null

$($global:CheatSheat.Threads.ActiveWindow.PowerShell).Runspace = $global:CheatSheat.Threads.ActiveWindow.Runspace
$global:CheatSheat.Threads.ActiveWindow.Handle = $($global:CheatSheat.Threads.ActiveWindow.PowerShell).BeginInvoke()

$Window_XAML = @'
<Window 
    Title="CheatSheetCompanion" SizeToContent="Manual" Topmost="True" 
    xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
    xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
    xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"
'@

if($global:CheatSheat.General.WindowType -eq 'WebView2') {
    $Window_XAML += '    xmlns:Wpf="clr-namespace:Microsoft.Web.WebView2.Wpf;assembly=Microsoft.Web.WebView2.Wpf"'
}

$Window_XAML += @'
    Width="400" Height="300" WindowStyle="ToolWindow" Opacity="0.85">
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
                <ComboBox Name="MarkdownComboBox" IsEditable="False" VerticalAlignment="Top" HorizontalAlignment="Left" Width="138" ToolTip="Select a note to view.">
'@

Get-ChildItem -Path $($global:CheatSheat.General.CurrentPath + '\md') -Filter '*.md' -Recurse | ForEach-Object {
    $Window_XAML += '                <ComboBoxItem>' + $($($_.BaseName  -replace '_','/') -replace '-','//') + '</ComboBoxItem>'
}

$Window_XAML += @'            
                </ComboBox>
            </StackPanel>    
            <StackPanel Grid.Column="1" Grid.Row="0" Orientation="Horizontal" Margin="5,15,5,0">
                <Button Background="LightGray" Name="RefreshButton" BorderBrush="Gray" BorderThickness="1" Height="22" Width="25" HorizontalAlignment="Right"  Margin="0,0,5,0" VerticalAlignment="Top" ToolTip="Refresh all notes.">
                    <StackPanel Orientation="Horizontal">
                        <TextBlock FontFamily="Segoe UI Symbol" Text="&#xE117;" Name="txtRefreshMarkdown" HorizontalAlignment="Right"/>
                    </StackPanel>
                </Button>
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
                        <TextBlock FontFamily="Segoe UI Symbol" Text="&#xE149;" Name="txtLastActiveWindow" HorizontalAlignment="Right"/>
                    </StackPanel>
                </Button>
            </StackPanel>    
'@            
if($global:CheatSheat.General.WindowType -eq 'WebView2') {
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

$global:CheatSheat.AutoSwitch.SynchronizeWindow = [Windows.Markup.XamlReader]::Parse("$strWrite")
$window = $global:CheatSheat.AutoSwitch.SynchronizeWindow

Remove-Variable -Name Window_XAML -Force

$Window.FindName('MinimizeButton').`
    add_Click({
        if($Window.Height -lt 15) {
            $Window.Height = $global:CheatSheat.General.WindowHeight

            $Window.FindName('InnerGrid').Visibility = 'Visible'
            $Window.FindName('StatusBar').Visibility = 'Visible'
            $Window.WindowStyle = [System.Windows.WindowStyle]::ToolWindow
            $Window.ResizeMode = [System.Windows.ResizeMode]::CanResize

            $Window.FindName('MinimizeText').Text = [System.Net.WebUtility]::HtmlDecode('&#x1f53a;&#x1f53a;&#x1f53a;') 
        }
        else {
            $global:CheatSheat.General.WindowHeight = $Window.Height

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
            $file = Get-Content -Path $($global:CheatSheat.General.CurrentPath + '\Output\' + $($($selectedItem.content  -replace '/','_') -replace '//','-') + '.html') -Raw

            if($null -ne $file) {
                $Window.FindName('WebBrowser').NavigateToString($(
                    '<html><head><meta charset="UTF-8"></head>' + `
                    '<link rel="stylesheet" href="./styles/default.min.css">' + `
                    '<script src="highlight.min.js"></script>' + `
                    '<script>hljs.highlightAll();</script>' + `
                    $(if(Test-Path $($global:CheatSheat.General.CurrentPath + '\styles\userstyle.css')){
                         '<style>' + $(Get-Content -Path $($global:CheatSheat.General.CurrentPath + '\styles\userstyle.css') -Raw) + '</style>'
                    })  + `
                    '<body class="markdown-body">' + $file + '</body></html>'
                ))
            }
        }
    })    

$Window.FindName('txtRefreshMarkdown').`
    Add_MouseDown({

        Get-ChildItem -Path $($global:CheatSheat.General.CurrentPath + '\Output') | Remove-Item -Force

        $MarkdownComboBox = $Window.FindName('MarkdownComboBox')
        $MarkdownComboBox.items.Clear()

        Get-ChildItem -Path $($global:CheatSheat.General.CurrentPath + '\md') -Filter '*.md' -Recurse | ForEach-Object {
            $ComboBoxItem1 = New-Object System.Windows.Controls.ComboBoxItem
            $ComboBoxItem1.Content = $($($_.BaseName -replace '_','/') -replace '-','//')
            $MarkdownComboBox.Items.Add($ComboBoxItem1)
        }

        $initialSessionState = [System.Management.Automation.Runspaces.InitialSessionState]::CreateDefault()

        $initialSessionState.Variables.Add(
                (New-object System.Management.Automation.Runspaces.SessionStateVariableEntry -ArgumentList "CheatSheat",$CheatSheat,$Null))

        $global:CheatSheat.Threads.MarkDownConvert.Runspace = [System.Management.Automation.Runspaces.RunspaceFactory]::CreateRunspace($Host, $initialSessionState)
        $($global:CheatSheat.Threads.MarkDownConvert.Runspace).ApartmentState = "STA"
        $($global:CheatSheat.Threads.MarkDownConvert.Runspace).ThreadOptions = "ReuseThread"
        $($global:CheatSheat.Threads.MarkDownConvert.Runspace).Open()


        $Code = {
            Get-ChildItem -Path $($global:CheatSheat.General.CurrentPath + '\md') -Filter '*.md' -Recurse | ForEach-Object {
                $(ConvertFrom-Markdown -Path $_.FullName).Html | Out-File $($global:CheatSheat.General.CurrentPath + '\Output\' + $_.BaseName + '.html') -Force
            }
        }

        $global:CheatSheat.Threads.MarkDownConvert.PowerShell = [PowerShell]::Create()
        $($global:CheatSheat.Threads.MarkDownConvert.PowerShell).AddScript($Code) | Out-Null
        
        $($global:CheatSheat.Threads.MarkDownConvert.PowerShell).Runspace = $global:CheatSheat.Threads.MarkDownConvert.Runspace
        $global:CheatSheat.Threads.MarkDownConvert.Handle = $($global:CheatSheat.Threads.MarkDownConvert.PowerShell).BeginInvoke()

        $($global:CheatSheat.Threads.MarkDownConvert.PowerShell).EndInvoke($global:CheatSheat.Threads.MarkDownConvert.Handle[0])
    })

$Window.FindName('txtEditMarkdown').`
    Add_MouseDown({    
        $selectedItem = $null
        $selectedItem = $Window.FindName('MarkdownComboBox').items | Where-Object {$_.IsSelected -eq $true}
        if($null -ne $selectedItem) {
            $file = $null
            $file = Get-ChildItem -Path $($global:CheatSheat.General.CurrentPath + '\md\') -Recurse -Filter $($($($selectedItem.content -replace '/','_') -replace '//','-') + '.md') | Select-Object -First 1
            if($null -ne $file) {
                Start-Process -FilePath $file.FullName               
            }
        }
        else {
            $Window.FindName('txtStatus').Text = 'Choose a markdown file first before trying to edit.'
        }
    })

$Window.FindName('txtAddMarkdown').`
    Add_MouseDown({    

        if($global:CheatSheat.AutoSwitch.Active -eq $true) {
            if(-not(Test-Path $($global:CheatSheat.General.CurrentPath + '\md\' + $global:CheatSheat.AutoSwitch.LastProcess.Process.Name + '.md'))) {
                New-Item -Path $($global:CheatSheat.General.CurrentPath + '\md') -Name $($global:CheatSheat.AutoSwitch.LastProcess.Process.Name + '.md')

                Start-Process -FilePath $($global:CheatSheat.General.CurrentPath + '\md\' + $global:CheatSheat.AutoSwitch.LastProcess.Process.Name + '.md')
                $Window.FindName('txtStatus').Text = 'Created a ' + $global:CheatSheat.AutoSwitch.LastProcess.Process.Name + ' markdown file.'
            }
        }
        else {
            Start-Process -FilePath $(New-Item -Path $($global:CheatSheat.General.CurrentPath + '\md') -Name $('PROCESS_OR_TITLE_' + $(Get-Random) + '.md') -ItemType File).FullName
            $Window.FindName('txtStatus').Text = 'Created a random markdown file.'
        }
    })

$Window.FindName('txtLastActiveWindow').`
    Add_MouseDown({
        if([System.Windows.Forms.Control]::MouseButtons -eq [System.Windows.Forms.MouseButtons]::Left) {
            if($global:CheatSheat.AutoSwitch.Active -eq $true) {
                $Window.FindName('txtStatus').Text = 'Deactivate AutoSwitch first'
            }
            elseif($null -ne $global:CheatSheat.LastActiveProcess) {
				
				$file = $null
				
				if(Test-Path $($global:CheatSheat.General.CurrentPath + '\Output\' + $global:CheatSheat.LastActiveProcess.Process.Name + '.html')){
					$file = Get-Content -Path $($global:CheatSheat.General.CurrentPath + '\Output\' + $global:CheatSheat.LastActiveProcess.Process.Name + '.html') -Raw
                }

                if($null -ne $file) {
                    $Window.FindName('WebBrowser').NavigateToString($('<html><head><meta charset="UTF-8"></head>' + $(if(Test-Path $($global:CheatSheat.General.CurrentPath + '\styles\userstyle.css')){ '<style>' + $(Get-Content -Path $($global:CheatSheat.General.CurrentPath + '\styles\userstyle.css') -Raw) + '</style>'})  + '<body class="markdown-body">' + $file + '</body></html>'))
                }
                else {
                    Get-ChildItem -Path $($global:CheatSheat.General.CurrentPath + '\md') | ForEach-Object {
                        if($global:CheatSheat.LastActiveProcess.Title -like $('*' + $_.BaseName + '*')) {
                            $file = Get-Content -Path $_.FullName -Raw
                            
                            $Window.FindName('WebBrowser').NavigateToString($('<html><head><meta charset="UTF-8"></head>' + $(if(Test-Path $($global:CheatSheat.General.CurrentPath + '\styles\userstyle.css')){ '<style>' + $(Get-Content -Path $($global:CheatSheat.General.CurrentPath + '\styles\userstyle.css') -Raw) + '</style>'})  + '<body class="markdown-body">' + $file + '</body></html>'))
                        }
                    }
                }
            }
        }
        elseif([System.Windows.Forms.Control]::MouseButtons -eq [System.Windows.Forms.MouseButtons]::Right) {
            if($global:CheatSheat.AutoSwitch.Active -eq $true) {
                $global:CheatSheat.AutoSwitch.Active = $false
                $Window.FindName('txtLastActiveWindow').Foreground = '#FF000000'
                $Window.FindName('txtStatus').Text = 'Auto Sync deactivated.'                
            }
            else {
                $global:CheatSheat.AutoSwitch.Active = $true
                $Window.FindName('txtLastActiveWindow').Foreground = '#ED1C24'
            }

            $global:CheatSheat.Threads.ActiveWindow.PowerShell.Stop()
            $global:CheatSheat.Threads.ActiveWindow.PowerShell.Streams.ClearStreams()
					
            $global:CheatSheat.Threads.ActiveWindow.Runspace.InitialSessionState.Variables.Remove("CheatSheat",$null)

            $global:CheatSheat.Threads.ActiveWindow.Runspace.InitialSessionState.Variables.Remove("Window",$null)
            
            $global:CheatSheat.Threads.ActiveWindow.Runspace.InitialSessionState.Variables.Add(
                (New-object System.Management.Automation.Runspaces.SessionStateVariableEntry -ArgumentList "CheatSheat",$global:CheatSheat,$Null)
            )

            $global:CheatSheat.Threads.ActiveWindow.Runspace.InitialSessionState.Variables.Add(
						(New-object System.Management.Automation.Runspaces.SessionStateVariableEntry -ArgumentList "ShowResult",$($global:CheatSheat.AutoSwitch.SynchronizeWindow.FindName("ShowResult")),$Null))

            $global:CheatSheat.Threads.ActiveWindow.Runspace.InitialSessionState.Variables.Add(
                (New-object System.Management.Automation.Runspaces.SessionStateVariableEntry -ArgumentList "Window",$Window,$Null)
            )

            $global:CheatSheat.Threads.ActiveWindow.PowerShell.Runspace.ResetRunspaceState()
            $global:CheatSheat.Threads.ActiveWindow.Handle = $($global:CheatSheat.Threads.ActiveWindow.PowerShell).BeginInvoke()						
        }
   })        

if($global:CheatSheat.General.WindowType -ne 'WebView2') {   
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
		(New-object System.Management.Automation.Runspaces.SessionStateVariableEntry -ArgumentList "CheatSheat",$global:CheatSheat,$Null))

    $initialSessionState.Variables.Add(
        (New-object System.Management.Automation.Runspaces.SessionStateVariableEntry -ArgumentList "Window",$Window,$Null))

    $global:CheatSheat.Threads.GlobalShortcut.Runspace = [System.Management.Automation.Runspaces.RunspaceFactory]::CreateRunspace($Host, $initialSessionState)
    $($global:CheatSheat.Threads.GlobalShortcut.Runspace).ApartmentState = "STA"
    $($global:CheatSheat.Threads.GlobalShortcut.Runspace).ThreadOptions = "ReuseThread"
    $($global:CheatSheat.Threads.GlobalShortcut.Runspace).Open()

    $Code = {
        while ($true) {
            if($(Invoke-Command $([System.Management.Automation.ScriptBlock]::Create('[CheatSheetCompanionHelper.KeyLogger]::WaitForKey(' + $($($($global:CheatSheat.General.INIConfig.GlobalShortcut.DisplayCompanion -split ',') | ForEach-Object {[System.Management.Automation.ScriptBlock]::Create($('[System.Windows.Forms.Keys]::' + $_))}) -join ',') +')')) -NoNewScope)) {
                $global:CheatSheat.AutoSwitch.SynchronizeWindow.Dispatcher.invoke(	[action]{
                    if($Window.Height -lt 15) {
                        $Window.Height = $global:CheatSheat.General.WindowHeight
                        
            
                        $Window.FindName('InnerGrid').Visibility = 'Visible'
                        $Window.FindName('StatusBar').Visibility = 'Visible'
                        $Window.WindowStyle = [System.Windows.WindowStyle]::ToolWindow
                        $Window.ResizeMode = [System.Windows.ResizeMode]::CanResize
            
                        $Window.FindName('MinimizeText').Text = [System.Net.WebUtility]::HtmlDecode('&#x1f53a;&#x1f53a;&#x1f53a;') 
                    }
                    else {
                        $global:CheatSheat.General.WindowHeight = $Window.Height
            
                        $Window.FindName('InnerGrid').Visibility = 'Hidden'
                        $Window.FindName('StatusBar').Visibility = 'Hidden'
                        $Window.FindName('MinimizeText').Text = [System.Net.WebUtility]::HtmlDecode('&#x1f53b;&#x1f53b;&#x1f53b;')            
                        $Window.WindowStyle = [System.Windows.WindowStyle]::None
                        $Window.ResizeMode = [System.Windows.ResizeMode]::NoResize
                        $Window.Height = '13'
                    }
                })
            }
            elseif($(Invoke-Command $([System.Management.Automation.ScriptBlock]::Create('[CheatSheetCompanionHelper.KeyLogger]::WaitForKey(' + $($($($global:CheatSheat.General.INIConfig.GlobalShortcut.AutoSwitch -split ',') | ForEach-Object {[System.Management.Automation.ScriptBlock]::Create($('[System.Windows.Forms.Keys]::' + $_))}) -join ',') +')')) -NoNewScope)) {
                $global:CheatSheat.AutoSwitch.SynchronizeWindow.Dispatcher.invoke(	[action]{
                    if($global:CheatSheat.AutoSwitch.Active -eq $true) {
                        $global:CheatSheat.AutoSwitch.Active = $false
                        $Window.FindName('txtLastActiveWindow').Foreground = '#FF000000'
                        
                        $Window.FindName('txtStatus').Text = 'Auto Sync deactivated.'
                        
                    }
                    else {
                        $global:CheatSheat.AutoSwitch.Active = $true
                        $Window.FindName('txtLastActiveWindow').Foreground = '#ED1C24'
                    }

                    $global:CheatSheat.Threads.ActiveWindow.PowerShell.Stop()
                    $global:CheatSheat.Threads.ActiveWindow.PowerShell.Streams.ClearStreams()
                            
                    $global:CheatSheat.Threads.ActiveWindow.Runspace.InitialSessionState.Variables.Remove("CheatSheat",$null)

                    $global:CheatSheat.Threads.ActiveWindow.Runspace.InitialSessionState.Variables.Remove("Window",$null)
                    
                    $global:CheatSheat.Threads.ActiveWindow.Runspace.InitialSessionState.Variables.Add(
                        (New-object System.Management.Automation.Runspaces.SessionStateVariableEntry -ArgumentList "CheatSheat",$global:CheatSheat,$Null)
                    )

                    $global:CheatSheat.Threads.ActiveWindow.Runspace.InitialSessionState.Variables.Add(
                                (New-object System.Management.Automation.Runspaces.SessionStateVariableEntry -ArgumentList "ShowResult",$($global:CheatSheat.AutoSwitch.SynchronizeWindow.FindName("ShowResult")),$Null))

                    $global:CheatSheat.Threads.ActiveWindow.Runspace.InitialSessionState.Variables.Add(
                        (New-object System.Management.Automation.Runspaces.SessionStateVariableEntry -ArgumentList "Window",$Window,$Null)
                    )

                    $global:CheatSheat.Threads.ActiveWindow.PowerShell.Runspace.ResetRunspaceState()
                    $global:CheatSheat.Threads.ActiveWindow.Handle = $($global:CheatSheat.Threads.ActiveWindow.PowerShell).BeginInvoke()
                
                })	
            }
        }
    }

    $global:CheatSheat.Threads.GlobalShortcut.PowerShell = [PowerShell]::Create()
    $($global:CheatSheat.Threads.GlobalShortcut.PowerShell).AddScript($Code) | Out-Null

    $($global:CheatSheat.Threads.GlobalShortcut.PowerShell).Runspace = $global:CheatSheat.Threads.GlobalShortcut.Runspace
    $global:CheatSheat.Threads.GlobalShortcut.Handle = $($global:CheatSheat.Threads.GlobalShortcut.PowerShell).BeginInvoke()

})


$window.Add_SourceInitialized({
    $window.Add_LocationChanged({
        $sensitivity = 35
        if ($window.Top -lt $sensitivity) {
            $window.Top = 0
        }
    })
})

if($global:CheatSheat.General.WindowType -eq 'WebView2') {
    $window.FindName("WebBrowser").CreationProperties = New-Object Microsoft.Web.WebView2.Wpf.CoreWebView2CreationProperties
    $window.FindName("WebBrowser").CreationProperties.UserDataFolder = "$env:LocalAppData\CheatSheetCompanion" 
}    

$window.ShowDialog()

$CheatSheat.Threads.ActiveWindow.PowerShell.Stop()
$CheatSheat.Threads.ActiveWindow.Handle = $null
$CheatSheat.Threads.ActiveWindow.PowerShell = $null
$CheatSheat.Threads.ActiveWindow.Runspace.Close()
$CheatSheat.Threads.ActiveWindow.Runspace.Dispose()
$CheatSheat.Threads.ActiveWindow.Runspace = $null


$CheatSheat.Threads.GlobalShortcut.PowerShell.Stop()
$CheatSheat.Threads.GlobalShortcut.Handle = $null
$CheatSheat.Threads.GlobalShortcut.PowerShell = $null
$CheatSheat.Threads.GlobalShortcut.Runspace.Close()
$CheatSheat.Threads.GlobalShortcut.Runspace.Dispose()
$CheatSheat.Threads.GlobalShortcut.Runspace = $null


if($null -ne $CheatSheat.Threads.MarkDownConvert.PowerShell) {
    $CheatSheat.Threads.MarkDownConvert.PowerShell.Stop()
    $CheatSheat.Threads.MarkDownConvert.Handle = $null
    $CheatSheat.Threads.MarkDownConvert.PowerShell = $null
    $CheatSheat.Threads.MarkDownConvert.Runspace.Close()
    $CheatSheat.Threads.MarkDownConvert.Runspace.Dispose()
    $CheatSheat.Threads.MarkDownConvert.Runspace = $null
}

[GC]::Collect()
[GC]::WaitForPendingFinalizers()
