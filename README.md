# CheatSheetCompanion_Develop

## About

Cheat Sheet Companion is a powerful PowerShell tool that provides context-sensitive cheat sheets to users. </br>
No matter what program or website the user is in, the appropriate cheat sheets are always displayed, </br>
making it easier for users to find the information they need quickly and efficiently.

## Workflow explanation

At the start of this program every markdown file will be converted to a HTML file with the correspondig name.
<b>When "AutoSwitch" is activated</b>, the program searches its library for relevant documents based on the active 
window title and process name. Each document's filename should contain at least one "tag" corresponding to a process 
name or a portion of the window title. Multiple tags can be separated by "-" or "_".
</br>
For example, a library file named "chrome_bing_firefox.md" would be displayed if any of the three tags (chrome, bing, firefox) 
appear in either the title description or the process name. Another example would be "git-pwsh.md" (git, pwsh).
</br>
If multiple HTML files are found that match the process or title, they will be added to a single output.
The content of the last written file will appear at the top of the display, followed by the rest of the content.
</br>
At the beginning of the program, each markdown file will be converted to an HTML file with the corresponding name.

## Requirements

Powershell 7 x64 - [link](https://github.com/PowerShell/PowerShell)

### Browser engines

Cheat Sheet Companion support the default system browser as well as the [WebView2](https://developer.microsoft.com/en-us/microsoft-edge/webview2/) engine if installed.</br>
To install the [WebView2](https://developer.microsoft.com/en-us/microsoft-edge/webview2/) engine</br>
run the following code as an administrator:

```pwsh
Import-Module PackageManagement
Install-Module PowerShellGet -AllowClobber -Force
If ((Get-PackageSource | Where Name -eq nuget.org) -eq $Null){
    Register-PackageSource -Name nuget.org -Location https://www.nuget.org/api/v2 -ProviderName NuGet -Trusted
}

Install-Package Microsoft.Web.WebView2 -Source nuget.org
```

## Customization

### INI file

> In this file, you can change the global shortcuts.

```
.\ini\config.ini
```

example:
```
[GlobalShortcut]
AutoSwitch = ControlKey,ShiftKey,F2
DisplayCompanion = ControlKey,ShiftKey,F11
```

### CSS file

> In this file, you can change the graphical appearance.

```
.\styles\userstyle.css
```

The HTML Code will be generated within a body entity with the class "markdown-body".

example:

```
<html>
	<head>
		<meta charset="UTF-8">
		</head>
		<body class="markdown-body">
			<h1 id="EXAMPLE_H1">EXAMPLE Header 1</h1>
			<h2 id="EXAMPLE_H2">EXAMPLE Header 2</h2>
			<p>EXAMPLE
			<table>
			<thead>
				<tr>
					<th>Symbol</th>
					<th>description</th>
				</tr>
			</thead>
			<tbody>
				<tr>
					<td>
						<p>EXAMPLE</p>
					</td>
					<td></td>
				</tr>
			</tbody>
		</table>
	</body>
</html>
```
