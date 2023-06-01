### SAPLauncher
Opens a new 
``` vbscript
Dim wsh : Set wsh = CreateObject("Wscript.Shell")
Set launcher = New SAPLauncher
launcher.SetClientName = 105
launcher.SetSystemName = "FQ2"
launcher.SetLocalXML = wsh.ExpandEnvironmentStrings("%APPDATA%") & "\SAP\Common\SAPUILandscape.xml"
launcher.CheckSAPLogon
launcher.FindSAPSession

If Not launcher.SessionFound Then 
	WScript.Quit(1)
End If

'Dosomething with the session object
 launcher.GetSession
```
