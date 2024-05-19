# CheatSheetCompanion_Develop

## About

Cheat Sheet Companion is a powerful PowerShell tool that provides context-sensitive cheat sheets to users.</br>
No matter what program or website the user is using, the appropriate cheat sheets are always displayed,</br>
making it easier for users to find the information they need quickly and efficiently.</br>
It also supports Obsidian and call directly files in the Obsidian vault.

## Workflow explanation

At the start of this program every markdown file will be converted to a HTML file with the correspondig name.</br>
After a few seconds the program searches for the title description of the active window and then the process name</br>
based on the file name similarity (Cosine Similarity).

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
