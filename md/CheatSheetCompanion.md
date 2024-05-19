# Manual

## About

Cheat Sheet Companion is a powerful PowerShell tool that provides context-sensitive cheat sheets to users.
No matter what program or website the user is using, the appropriate cheat sheets are always displayed,
making it easier for users to find the information they need quickly and efficiently.
It also supports Obsidian and call directly files in the Obsidian vault.

## GUI legend

> Cheat Sheet Companion has a magic pinning to the top window if the application window is near.

| Symbol                                                                   | description                                                                                              |
| ------------------------------------------------------------------------ | -------------------------------------------------------------------------------------------------------- |
| <p style="font-family: Segoe UI Symbol;">&#xE104;</p>                    | Edit the currently displayed Markdown file in the default system editor or Obsidian.                     |
| <p style="font-family: Segoe UI Symbol;">&#xE109;</p>                    | Create a Markdown file for the current process and open it in the default system editor or Obsidian.     |
| <p style="font-family: Segoe UI Symbol;">&#xE149;</p>                    | Click to activate or deactivate "AutoSwitch".                                                            |
| <p style="font-family: Segoe UI Symbol;">&#x1f53a;&#x1f53a;&#x1f53a;</p> | Minimize the Cheat Sheet Companion. </br>With a <b>right click</b>, close this application.              |
| Combobox                                                                 | Use this to switch between the different generated HTML files manually.                                  |
| Statusbar                                                                | When "AutoSwitch" is activated, it will display the last active window's current title and process name. |


## Workflow explanation

At the start of this program every markdown file will be converted to a HTML file with the correspondig name.
After a few seconds the program searches for the title description of the active window and then the process name
based on the file name similarity (Cosine Similarity).

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



