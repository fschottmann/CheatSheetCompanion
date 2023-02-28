# Manual

## About

Cheat Sheet Companion is a powerful PowerShell tool that provides context-sensitive cheat sheets to users.
No matter what program or website the user is in, the appropriate cheat sheets are always displayed,
making it easier for users to find the information they need quickly and efficiently.

## GUI legend

> Cheat Sheet Companion has a magic pinning to the top window if the application window is near.

| Symbol                                                                   | description                                                                                                                                                                                  |
| ------------------------------------------------------------------------ | -------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- |
| <p style="font-family: Segoe UI Symbol;">&#xE117;</p>                    | Refresh all HTML pages according to the actual markdown file.                                                                                                                                |
| <p style="font-family: Segoe UI Symbol;">&#xE104;</p>                    | Edit in the combo box selected Markdown file in the default system editor.                                                                                                                   |
| <p style="font-family: Segoe UI Symbol;">&#xE109;</p>                    | Create a new markdown file if "AutoSwitch" is enabled;</br>it creates a markdown file for the current process.                                                                               |
| <p style="font-family: Segoe UI Symbol;">&#xE149;</p>                    | Grab the last active window and open the corresponding HTML file.</br>With a <b>right click</b>, the "AutoSwitch" will be activated.</br>"AutoSwitch" display the corresponding HTML automatically. |
| <p style="font-family: Segoe UI Symbol;">&#x1f53a;&#x1f53a;&#x1f53a;</p> | Minimize the Cheat Sheet Companion. </br>With a <b>right click</b>, close this application.                                                                                                                                                         |
| Combobox                                                                 | Use this to switch between the different generated HTML files manually.                                                                                                                      |
| Statusbar                                                                | When "AutoSwitch" is activated, it will display the last active window's current title and process name.                                                                                     |


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

## Recommended usage

To make the best use of Cheat Sheet Companion, one could create a task scheduler entry.

Example:

```pwsh
Programm: "C:\Program Files\PowerShell\7\pwsh.exe" 
Arguments: -NoProfile -NonInteractive -WindowStyle Hidden C:\Path\To\CheatSheetCompanion\CheatSheetCompanion.ps1
```


## Browser engines

Cheat Sheet Companion supports the default system browser as well as the [WebView2](https://developer.microsoft.com/en-us/microsoft-edge/webview2/) engine if installed.</br>
To install the [WebView2](https://developer.microsoft.com/en-us/microsoft-edge/webview2/) engine
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



