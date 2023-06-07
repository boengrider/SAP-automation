### SAPLauncher
* Opens a new connection to the specified system and client and returns a session object  
* Class supports multiclient

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

### GetExcelWorkbook  
Sometimes when interacting with SAP via scripting engine, documents are exported in a bit different way.  
Function simply searches for the desired workbooks.
However it only searches 1st instance of an excel process

``` vbscript
Dim wb
If GetExcelWorkbook("report.xlsx",wb) Then 
  'Found
Else
  'Not found
End If
```
