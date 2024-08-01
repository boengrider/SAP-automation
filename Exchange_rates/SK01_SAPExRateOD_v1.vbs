Option Explicit 

Dim oDF,oTCD,oWSH,FSO,t,m,system,client,oSL,oRU
Dim args,arg,choice,ss
Dim dictValidSystems,dictValidClients
Dim oRx
Dim boolInteractive : boolInteractive = False 

Set dictValidClients = CreateObject("Scripting.Dictionary")
Set dictValidSystems = CreateObject("Scripting.Dictionary")
Set args = CreateObject("Scripting.Dictionary")

Set FSO = CreateObject("Scripting.FileSystemObject")
Set oRU = New RateUpload_v1
Set oSL = New SAPLauncher
Set oRx = New RegExp
Set t = New ECBRate
Set oWSH = CreateObject("Wscript.Shell")
Set oDF = New DateFormatter
Set oTCD = New TCDCalendar

dictValidSystems.Add "fq2",""
dictValidSystems.Add "fp2",""
dictValidClients.Add "105",""
oRx.Pattern = "[0-9]{2}\.[0-9]{2}\.[0-9]{4}"
'********************* M A I N ********************************

'************ Check if we're running in Wscript ***************
If InStr(LCase(WScript.FullName),"wscript.exe") > 0 Then
	boolInteractive = True
	'Prompt for input
	choice = InputBox("Enter the date for which you wish to upload the exchange rates" & vbCr & vbCr & "Example: 24.12.2020","Enter the target date")
	If choice = "" Then
		MsgBox "Program cannot continue without a date",vbOKOnly + vbCritical,"ERROR"
		WScript.Quit
	End If 
	
	Do While Not oRx.Test(choice)
		MsgBox "Invalid date format" & vbCr & vbCr & "Enter a date like this: 24.12.2020 or 01.05.2020",vbOK,"Warning: Invalid date format"
		choice = InputBox("Enter the date for which you wish to upload the exchange rates" & vbCr & vbCr & "Example: 24.12.2020","Enter the target date")
	Loop 
	choice = Replace(choice,".","/") ' Replace '.' with '/'
	args.Add "date",CDate(choice)
	
	choice = InputBox("0 -> Download rates only" & vbCr & "1 -> Download & Upload to SAP","Select action","0")
	If choice = "" Then ' User pressed cancel
		WScript.Quit
	Else
		Select Case choice
		
			Case "0"
				args.Add "downloadonly",True
				
			Case "1"
				args.Add "downloadonly",False
				
			Case Else
				MsgBox "Unrecognized option",vbOK + vbCritical,"ERROR"
				WScript.Quit
		End Select
	End If 
	
	If Not args.Item("downloadonly") then
		choice = InputBox("Enter the target SAP system" & vbCr & vbCr & "Example: FQ2","Select the target SAP system","FQ2")
		If choice = "" Then
			MsgBox "Program cannot continue without a SAP system name",vbOKOnly + vbCritical,"ERROR"
			WScript.Quit
		End If 
		If Not dictValidSystems.Exists(LCase(choice)) Then
			sgBox "Uknown system name",vbOKOnly + vbCritical,"Uknown SAP system"
			WScript.Quit
		End If
		args.Add "system", LCase(choice)
		
		choice = InputBox("Enter the target client number" & vbCr & vbCr & "Example: 105","Select the targert SAP client","105")
		If choice = "" Then
			MsgBox "Program cannot continue without a SAP client number",vbOKOnly + vbCritical,"ERROR"
			WScript.Quit
		End If 
		If Not dictValidClients.Exists(choice) Then
			MsgBox "Uknown client number",vbOKOnly + vbCritical,"Uknown SAP client"
			WScript.Quit
		End If
		args.Add "client", choice
	End If 
End If 
	 
'************ Process command line arguments ******************
If WScript.Arguments.Count > 0 Then 
	For arg = 0 To WScript.Arguments.Count - 1
		
		Select Case WScript.Arguments.Item(arg)
		
			case "--system","-s"
				If arg < WScript.Arguments.Count - 1 Then
					If args.Exists("system") Then
						args.Remove("system")
						args.Add "system",WScript.Arguments.Item(arg + 1)
					Else
						args.Add "system",WScript.Arguments.Item(arg + 1)
					End If
					If Not dictValidSystems.Exists(LCase(args.Item("system"))) Then
						WScript.Echo "Uknown SAP system name -> " & args.Item("system")
						WScript.Quit
					End If 
				End If
				
			case "--client","-c"
				If arg < WScript.Arguments.Count - 1 Then
					If args.Exists("client") Then
						args.Remove("client")
						args.Add "client",WScript.Arguments.Item(arg + 1)
					Else
						args.Add "client",WScript.Arguments.Item(arg + 1)
					End If
					If Not dictValidClients.Exists(args.Item("client")) Then
						WScript.Echo "Unknown SAP client number -> " & args.Item("client")
						WScript.Quit
					End If 
				End If 
				
			case "--date","-d"
				If arg < WScript.Arguments.Count - 1 Then
					If args.Exists("date") Then
						args.Remove("date")
						args.Add "date",WScript.Arguments.Item(arg + 1)
					Else
'						If Not oRx.Test(args.Item("date")) Then
'							WScript.Echo "Invalid date format. Enter a date like this: 24.12.2020 or 05.09.2020"
'							WScript.Echo CStr(oRx.Test(args.Item("date")))
'							WScript.Echo args.Item("date")
'							WScript.Quit
'						End If 
						args.Add "date",CDate(Replace(WScript.Arguments.Item(arg + 1),".","/"))
					End If
				End If
				
		End Select
	Next
End If 

If Not boolInteractive Then
	WScript.Echo "System -> " & args.Item("system")
	WScript.Echo "Client -> " & args.Item("client")
	WScript.Echo "Date   -> " & args.Item("date")
End If 		 
system = args.Item("system")
client = args.Item("client")
 		

' Initialize a TCDCalendar object
oTCD.AddTCD "01012020","New Year's Day"
oTCD.AddTCD "10042020","Good Friday"
oTCD.AddTCD "13042020","Easter Monday"
oTCD.AddTCD "01052020","Labour Day"
oTCD.AddTCD "25122020","Christmas Day"
oTCD.AddTCD "26122020","Christmas Holiday"
oTCD.AddTCD "01012021","New Year's Day"
oTCD.AddTCD "02042021","Good Friday"
oTCD.AddTCD "05042021","Easter Monday"
oTCD.AddTCD "01052021","Labour Day"
oTCD.AddTCD "25122021","Christmas Day"
oTCD.AddTCD "26122021","Christmas Holiday"
oTCD.AddTCD "01012022","New Year's Day"
oTCD.AddTCD "15042022","Good Friday"
oTCD.AddTCD "18042022","Easter Monday"
oTCD.AddTCD "01052022","Labour Day"
oTCD.AddTCD "25122022","Christmas Day"
oTCD.AddTCD "26122022","Christmas Holiday"
oTCD.AddTCD "01012023","New Year's Day"
oTCD.AddTCD "07042023","Good Friday"
oTCD.AddTCD "10042023","Easter Monday"
oTCD.AddTCD "01052023","Labour Day"
oTCD.AddTCD "25122023","Christmas Day"
oTCD.AddTCD "26122023","Christmas Holiday"

If Not FSO.FolderExists(oWSH.ExpandEnvironmentStrings("%SYSTEMDRIVE%") & "\ExRate") Then
	FSO.CreateFolder oWSH.ExpandEnvironmentStrings("%SYSTEMDRIVE%") & "\ExRate"
End If
If Not FSO.FolderExists(oWSH.ExpandEnvironmentStrings("%SYSTEMDRIVE%") & "\ExRate\SK01") Then
	FSO.CreateFolder oWSH.ExpandEnvironmentStrings("%SYSTEMDRIVE%") & "\ExRate\SK01"
End If 
t.Init Null,"C:\ExRate\SK01"														' Initialize a new ECBRate instance
't.SetOutputFile = Year(args.Item("date")) & Month(args.Item("date")) & Day(args.Item("date")) & ".txt"					    	' Set output file name
t.SetOutputFile = Right("0000" & Year(args.Item("date")),4) & Right("00" & Month(args.Item("date")),2) & Right("00" & Day(args.Item("date")),2) & ".txt"
t.SetXmlURL = "https://www.ecb.europa.eu/stats/eurofxref/eurofxref-hist-90d.xml"	' Set xml url
t.AddTargetCurrency = "USD,JPY,CZK,DKK,GBP,HUF,PLN,SEK,CHF"
t.OverrideQuantity "CZK,HUF"

' --------------------------  M A I N --------------------------------

oTCD.FindNonTCDDate(args.Item("date"))	' Find the first non TCD date
Do While t.FoundDateInXML = False 
	t.MakeOutputFile True,oDF.ToYearMonthDayWithDashes(oTCD.FirstNonTCD)
	oTCD.FindNonTCDDate args.Item("date") - 1
Loop 

If boolInteractive = True And args.Item("downloadonly") = True Then
	MsgBox "C:\ExRate\SK01\" & Right("0000" & Year(args.Item("date")),4) & Right("00" & Month(args.Item("date")),2) & Right("00" & Day(args.Item("date")),2) & ".txt created"
ElseIf boolInteractive = False And args.Item("downloadonly") = True Then
	WScript.Echo "C:\ExRate\SK01\" & Right("0000" & Year(args.Item("date")),4) & Right("00" & Month(args.Item("date")),2) & Right("00" & Day(args.Item("date")),2) & ".txt created"
End If 

If args.Item("downloadonly") Then
	WScript.Quit
End If 

' Upload to SAP
Set oSL = New SAPLauncher
Set oRU = New RateUpload_v1
oSL.SetClientName = args.Item("client")
oSL.SetSystemName = args.Item("system")
oSL.SetLocalXML = oWSH.ExpandEnvironmentStrings("%APPDATA%") & "\SAP\Common\SAPUILandscape.xml"
oSL.CheckSAPLogon
oSL.FindSAPSession
Set ss = oSL.GetSession

If Not oSL.SAPSessionExists Then 
	WScript.Quit
End If 

oRU.SAPSession = ss
oRU.UploadRates "C:\ExRate\SK01\" & t.GetOutputFile,"ysk1"
	





'------------------------ E N D    M A I N ----------------------------------



'----------------------- F U N C T I O N S ----------------------------------



Function ClearIECache
	Dim shell
	Set shell = CreateObject("Wscript.Shell")
	shell.Run "RunDLL32.exe InetCpl.cpl,ClearMyTracksByProcess 8",0,True
End Function 

'------------------ C L A S S E S -----------------------------------------






'-------------------- ECBRate Class ---------------------------------------
Class ECBRate 

	Private strXmlUrl
	Private strNameSpace
	Private dictQuantity        
	Private dictFcurrs					' Dictionary hodling target currencies
	Private dictOutputFiles				' Dictionary holding processed output file paths i.e C:\ExRate\SK01\20200101.txt C:\ExRate\SK01\20200102.txt ...
	Private strOutputDir
	Private strOutputFile
	Private strTcurr
	Private boolFoundDateInXml
	Private oOutFile
	Private oWSH
	Private oXML
	Private oHTTP
	Private oFSO
	Private errno

	
	' Constructor and destructor
	Private Sub Class_Initialize
	
		Set dictFcurrs = CreateObject("Scripting.Dictionary")
		Set oFSO = CreateObject("Scripting.FileSystemObject")
		Set oHTTP = CreateObject("MSXML2.XMLHTTP")
		Set oXML = CreateObject("MSXML2.DOMDocument")
		Set dictQuantity = CreateObject("Scripting.Dictionary")
		Set oWSH = CreateObject("Wscript.Shell")
		
		boolFoundDateInXml = False
		strTcurr = "EUR"
		strNameSpace = "xmlns:gesmes='http://www.gesmes.org/xml/2002-08-01' xmlns='http://www.ecb.int/vocabulary/2002-08-01/eurofxref'"
		strXmlUrl = ""
		strOutputDir = ""
		
	End Sub 
	
	Private Sub Class_Terminate
	
	End Sub 
	
	' Public methods
	
	
	' Init()
	Public Function Init(strUrl,strOutDir)
	
		strOutputDir = strOutDir
		strXmlUrl = strUrl
		
	End Function
	
	' OverrideQuantity()
	Public Function OverrideQuantity(strCurrs)
	
		Dim curr
		For Each curr In Split(strCurrs,",")
		
			dictQuantity(curr) = "100"
			
		Next
		
	End Function 
	
	' MakeOutputFile()
	Public Function MakeOutputFile(boolClearIECache,strDate)
	
		Dim ChildNodes,ChildNode,Attributes,Attribute,i,key,delim
	
		If boolClearIECache Then
		
			ClearIE					' Clear internet explorer cache first
			
		End If 
		
		
		' Open http connection
		
		oHTTP.open "GET",strXmlUrl,False
		oHTTP.send
		
		If oHTTP.status <> 200 Then
			
			errno = 2        		' Set errno
			MakeOutputFile = -1 	' ERROR downloading rate file
			Exit Function
			
		End If 
		
		' Http request OK, continue loading xml
		oXML.load oHTTP.responseXML
		oXML.setProperty "SelectionNamespaces", strNameSpace
		Set ChildNodes = oXML.getElementsByTagName("Cube")
		
		i = 0
		
		Do While ChildNodes.item(i).attributes.length <> 1
		
			i = i + 1
			
		Loop
		
		
		On Error Resume next
		Do While IsObject(ChildNodes.item(i))
		
			debug.WriteLine ChildNodes.item(i).attributes.getNamedItem("time").text
			
			If strDate = ChildNodes.item(i).attributes.getNamedItem("time").text Then
				boolFoundDateInXml = True
				debug.WriteLine "Date found in the XML"
				Exit Do 				' Exit loop
			End If 
			
			i = i + (ChildNodes.item(i).childNodes.length) + 1
			
			If Not ChildNodes.item(i).hasChildNodes Then
			
				boolFoundDateInXml = False
				Exit Do 
				
			End If 
			
		Loop
		
		If Not boolFoundDateInXml Then
		
			errno = 3
			MakeOutputFile = -1
			Exit Function 
			
		End If 
		On Error GoTo 0 
		
		' Found the target date 
		' Select rates
		Set ChildNode = ChildNodes.item(i)
		Set ChildNodes = ChildNode.childNodes
		
		For Each ChildNode In ChildNodes 
		
			If dictFcurrs.Exists(ChildNode.attributes.getNamedItem("currency").text) Then
			
				dictFcurrs(ChildNode.attributes.getNamedItem("currency").text) = ChildNode.attributes.getNamedItem("rate").text
				
			End If 
			
		Next 
	
		' Dictionary hold pairs CURRENCY RATE
	
		If InStr(1 / 2,",") >= 1 Then
			delim = ","
		Else 
			delim = "."
		End If 
		
		For Each key In dictFcurrs.Keys
	
			dictFcurrs(key) = Replace(dictFcurrs.Item(key),".",delim) ' replace . with whatever is the system delimiter
			
		Next
		 
		Set oOutFile = oFSO.OpenTextFile(strOutputDir & "\" & strOutputFile,2,True)
		
		For Each key In dictFcurrs.Keys
		
			oOutFile.WriteLine key & vbTab & strTcurr & vbTab & FormatNumber(Round(1/dictFcurrs.Item(key),5) * dictQuantity.Item(key),5) & vbTab & dictQuantity.Item(key) & vbTab & "1"
		Next 
		
		MakeOutputFile = 0
		'CopyOutputFile
		
		
	End Function
	
	' CopytOutputFile()
	Private Sub CopyOutputFile
	
		oFSO.CopyFile strOutputDir & "\" & strOutputFile, "\\vcn.ds.volvo.net\cli-sd\sd1294\046629\output\01_SK01_ExRateProcessing\SK01\" & strOutputFile
	
	End Sub 
	
	Private Sub ClearIE
	
		oWSH.Run "RunDLL32.exe InetCpl.cpl,ClearMyTracksByProcess 8",0,True
		
	End Sub 
		
	
	' Properties
	Public Property Get FoundDateInXML
	
		FoundDateInXML = boolFoundDateInXml
		
	End Property 
	
	Public Property Let SetXmlURL(url)
	
		strXmlUrl = url
		
	End Property  
	
	Public Property Let SetOutputDirectory(dir)
	
		strOutputDir = dir
		
	End Property 
	
	Public Property Let SetOutputFile(file)
	
		strOutputFile = file
		
	End Property 
	
	Public Property Get GetOutputFile
	
		GetOutputFile = strOutputFile
		
	End Property 
	
	Public Property Get GetOutputDirectory
	
		GetOutputDirectory = strOutputDir
		
	End Property  
	
	Public Property Get GetXmlURL
	
		GetXmlURL = strXmlUrl
		
	End Property 
	
	Public Property Get GetErrorCode
	
		GetErrorCode = errno
		
	End Property 
	
	Public Property Let AddTargetCurrency(currs)
	
		Dim curr
		For Each curr In Split(currs,",") 
		
			dictFcurrs.Add curr,""
			dictQuantity.Add curr,"1"
			
		Next
		
	End Property
	
	
End Class 
	
		
	





' TCDCalendar Class
Class TCDCalendar
	' ------------- Private members ----------- 
	Private t_dict ' Scripting.Dictionary that holds TCD entries
	Private t_len ' t_dict.Count
	Private t_isDateTCD ' this variable is set to True if current item in the dictionary is a TCD. Usefull during iterations through the dictionary
	Private t_FirstNonTCDDay
	
	' ------------- Constructor ----------------------
	
	Private Sub Class_Initialize
		Set t_dict = CreateObject("Scripting.Dictionary")
		t_len = t_dict.Count
		t_FirstNonTCDDay = Null
		t_isDateTCD = False ' Initially false. GetLastNonTCDDate(D) uses this variable
	End Sub 
		
	Private Sub Class_Terminate
	
	End Sub 
	' ------------ Instance methods ------------------
	Public Function AddTCD(ddmmyyyy,holiday_name)
		t_dict.Add ddmmyyyy,holiday_name
		t_len = t_dict.Count
	End Function 
	
	Public Function FindNonTCDDate(D) ' Argument is a Date Object
		If t_dict.Exists(Right("00" & Day(D),2) & Right("00" & Month(D),2) & Right("0000" & Year(D),4)) Or Weekday(D) = 1 Or Weekday(D) = 7 Then 
			FindNonTCDDate = FindNonTCDDate((D - 1)) ' Recursive call to the GetLastNonTCDDate function
		Else
			t_FirstNonTCDDay = D
			Exit Function
		End If
	End Function 
	
	Public Function IsTodayTCD
		Dim key
		If Weekday(Date()) = 1 Or Weekday(Date()) = 7 Then 
			IsTodayTCD = True
			Exit Function
		End If 
		
		For Each key In t_dict.Keys
				If key = (Right("00" & Day(Date()),2) & Right("00" & Month(Date()),2) & Right("0000" & Year(Date()),4)) Then 
				IsTodayTCD = True
				Exit Function
			End If 
		Next
		
		IsTodayTCD = False
	End Function
	
		
	' ----------- Getters and Setters ---------------- 
	Public Property Get Len
		Len = t_len
	End Property 
	
	Public Property Get FirstNonTCD
		FirstNonTCD = t_FirstNonTCDDay
	End Property 
	
	
End Class




' DateFormatter Class
Class DateFormatter
	' Convert from YYYY-MM-DD to DD.MM.YYYY
	Public Function FromYyyyMmDd_WithDashesTo_DdMmYyyy_WithDots(strDate)
		Dim temp
		temp = Right(strDate,2) & "." ' Day
		temp = temp & Mid(strDate,6,2) & "." ' Month
		temp = temp & Left(strDate,4) ' Year
		FromYyyyMmDd_WithDashesTo_DdMmYyyy_WithDots = temp
	End Function 
	
	Public Function ToYearMonthDay(D) ' Date object
		ToYearMonthDay = Right("0000" & Year(D),4) & Right("00" & Month(D),2) & Right("00" & Day(D),2)
	End Function
	
	Public Function ToYearMonthDayWithDashes(D)
		ToYearMonthDayWithDashes = Right("0000" & Year(D),4) & "-" & Right("00" & Month(D),2) & "-" & Right("00" & Day(D),2)
	End Function
	
	Public Function ToDayMonthYearWithDots(D)
		ToDayMonthYearWithDots = Right("00" & Day(D),2) & "." & Right("00" & Month(D),2) & "." & Right("0000" & Year(D),4)
	End Function
	
	
End Class 




'------------- SAPLauncher --------------
Class SAPLauncher
	
	Private oNET
	Private oXML
	Private oWSH
	Private oFSO
	Private oSAPGUI
	Private oAPP
	Private oCON
	Private oSES
	Private strGlobalURL
	Private strLocalLandscapePATH
	Private boolSAPRunning  		' Indicates whether SPA Logon is runniied files
	Private boolSessionFound		' Set to true if session was found or created
	Private strSSN 					' Sap System Name e.g FQ2
	Private strSCN  	    		' Sap Client Name e.g. 105
	Private strSSD          		' Sap System Description. This string is found in the local landscape xml and used to connect to the sap system


	
	
	
	' ============== Constructor & Destructor ==================
	Private Sub Class_Initialize
	
		Set oNET = CreateObject("Wscript.NEtwork")
		Set oXML = CreateObject("MSXML2.DOMDocument")
		Set oFSO = CreateObject("scripting.filesystemobject")
		Set oWSH = CreateObject("wscript.shell")
		oSAPGUI = Null
		oAPP = Null
		oCON = Null
		oSES = Null 
		strSSN = Null
		strSCN = Null
		strGlobalURL = Null
		strLocalLandscapePATH = Null
		strSSD = Null
		boolSAPRunning = False
		boolSessionFound = False
		

	End Sub
	
	Private Sub Class_Terminate
		
	End Sub
	
	' ============ P U B L I C  &  P R I V A T E   M E T H O D S  &   S U B R O U T I N E S ===========
	
	

	' ---------- CheckSAPLogon
	Public Sub CheckSAPLogon
		Dim waitPeriod,waitTurns,currentTurn
		waitPeriod = 500 ' miliseconds
		waitTurns = 5 
		currentTurn = 1
		
		On Error Resume Next
		
		Set oSAPGUI = GetObject("SAPGUI") ' This fails is saplogon is not running. We're connecting (coCreating ? ) to the COM object not creating our own instance in this approach
		If err.number <> 0 And Not IsObject(oSAPGUI) Then
			debug.WriteLine "SAP logon is not running"
			oWSH.Run "saplogon.exe",2,False
			Set oSAPGUI = GetObject("SAPGUI")
			Do While Not IsObject(oSAPGUI) And currentTurn <= waitTurns
				debug.WriteLine "Waiting for sap logon"
				WScript.Sleep waitPeriod * currentTurn '1st time wait 500 ms, 2nd time wait 1000 ms etc...
				Set oSAPGUI = GetObject("SAPGUI")
				currentTurn = currentTurn + 1
			Loop
		End If
		
		On Error GoTo 0
		
		If Not IsObject(oSAPGUI) Then
			boolSAPRunning = False
			Exit Sub
		End If 	
		
		boolSAPRunning = True
	
	End Sub 
	
	
	
	' ---------- FindSAPSession
	Public Sub FindSAPSession
		Dim waitPeriod,waitTurns,currentTurn
		waitPeriod = 5000 ' miliseconds
		waitTurns = 5 ' 5 x 5000 = 20000 ms / 20 s
		currentTurn = 1
		
		If Not boolSAPRunning Then
			oSES = Null
			Exit Sub 
		End If  
		
		FindSAPSystemDescription
		
		If IsNull(strSSD) Then
			oSES = Null
			Exit Sub
		End If
		 
		Set oAPP = oSAPGUI.GetScriptingEngine
	
			
		For currentTurn = 1 To waitTurns
			Set oCON = oAPP.OpenConnection(strSSD,True,False) ' Open a new connection synchronously
			On Error Resume Next
			Set oSES = oCON.Children(0) ' Attach to the first session
			
			KillPopups(oSES)
			On Error GoTo 0
			If Not oSES.ActiveWindow.FindByName("sbar", "GuiStatusbar") Is Nothing Then
				 If InStr(oSES.ActiveWindow.FindByName("sbar", "GuiStatusbar").text, "Enter a valid SAP user or choose one from the list") > 0 Then
					oSES.ActiveWindow.findById("usr/txtRSYST-MANDT").text = strSCN
					oSES.ActiveWindow.findById("usr/txtRSYST-BNAME").text = oNET.UserName
					oSES.ActiveWindow.findById("usr/txtRSYST-LANGU").text = "EN"
					oSES.ActiveWindow.sendvkey.0
					KillPopups(oSES) ' In case of multiple connections
				End If 
			End If 
			KillPopups(oSES)
				
			If Not IsObject(oSES) Or IsNull(oSES) Then
				Debug.WriteLine "No session found, waiting " & currentTurn & " out of " & waitTurns & " turns"
				WScript.Sleep waitPeriod
			ElseIf IsObject(oSES) And IsNull(oSES) Then
				Debug.WriteLine "No session found, waiting " & currentTurn & " out of " & waitTurns & " turns"
				WScript.Sleep waitPeriod
			Else
				Exit For
			End If 	
		Next
		
		If Not IsObject(oSES) Or IsNull(oSES) Then
				Debug.WriteLine "No session found after 5 retries"
				boolSessionFound = False
				Exit Sub
		End If 
		
		If IsObject(oSES) And Not IsNull(oSES) Then
			If InStr(oSES.findById("wnd[0]/sbar/pane[0]").text,"No user exists") > 0 Then
				oCON.CloseConnection
				boolSessionFound = False 
				debug.WriteLine "Session found: " & boolSessionFound
				Exit Sub 
			End If
			
			boolSessionFound = True
			If InStr(oSES.ActiveWindow.FindByName("sbar", "GuiStatusbar").text, "Enter a valid SAP user or choose one from the list") > 0 Then
				oSES.ActiveWindow.findById("usr/txtRSYST-MANDT").text = strSCN
				oSES.ActiveWindow.findById("usr/txtRSYST-BNAME").text = oNET.UserName
				oSES.ActiveWindow.findById("usr/txtRSYST-LANGU").text = "EN"
				oSES.ActiveWindow.sendvkey.0
				KillPopups(oSES) ' In case of multiple connections
			End If 
			debug.WriteLine "Session found: " & boolSessionFound
			Exit Sub 
		End If
		
		oCON.CloseConnection
		boolSessionFound = False
		debug.WriteLine "Session found: " & boolSessionFound
			
	End Sub 

	
	
	' --------- FindSAPSystemDescription
	Private Sub FindSAPSystemDescription
	

		oXML.load(strLocalLandscapePATH) ' Locally stored XML
		oXML.setProperty "SelectionLanguage", "XPath"
		
		On Error Resume next
		strSSD =  oXML.selectSingleNode("//Landscape/Services/Service[starts-with(@name, '" & strSSN & "')]").attributes.getNamedItem("name").text
		On Error GoTo 0
		
		If Not IsNull(strSSD) then
			CheckSAPLogon
			Exit Sub  
		End If 
				
		strSSD = Null ' Not found
		
	End Sub 
	
	Public Function GetSession
	
		If IsNull(oSES) Or Not IsObject(oSES) Then
			GetSession = Null
		else
			Set GetSession = oSES
		End If 

	End Function
	
	Public Function KillPopups(ByRef objSession)
		Do While objSession.Children.Count > 1
			If InStr(objSession.ActiveWindow.Text, "System Message") > 0 Then
				objSession.ActiveWindow.sendVKey 12
			ElseIf InStr(objSession.ActiveWindow.Text, "Information") > 0 And InStr(objSession.ActiveWindow.PopupDialogText, "Exchange rate adjusted to system settings") > 0 Then
				objSession.ActiveWindow.sendVKey 0
			ElseIf InStr(objSession.ActiveWindow.Text, "Copyright") > 0 Then
				objSession.ActiveWindow.sendVKey 0
			ElseIf InStr(objSession.ActiveWindow.Text, "License Information for Multiple Logon") > 0 Then
				objSession.ActiveWindow.findById("usr/radMULTI_LOGON_OPT2").select
				objSession.ActiveWindow.sendVKey 0
			'ElseIF   'Insert next type of popup windows which you want to kill
			Else
				Exit Do
			End If
		Loop
	End Function 

	' ================= P R O P E R T I E S ====================
	Public Property Get SAPLogonRunning
		SAPLogonRunning = boolSAPRunning
	End Property 	
		
	Public Property Get SAPSessionExists
		If boolSAPRunning And Not IsNull(oSES) Then
			SAPSessionExists = True
		Else 
			SAPSessionExists = False
		End If
	End Property 
	
	Public Property Get SAPsysName
		SAPsysName = strSSN
	End Property 
	
	Public Property Get SAPcliName
		SAPcliName = strSCN
	End Property 
	
	Public Property Get LandscapeURL
		LandscapeURL = strGlobalURL
	End Property 
	
	
	Public Property Get SAPsysDescription
		SAPsysDescription = strSSD
	End Property 
	
	Public Property Get GetGlobalURL
	
		GetGlobalURL = strGlobalURL
		
	End Property 
	
	Public Property Let SetGlobalURL(url)
	
		strGlobalURL = url
		
	End Property 
	
	Public Property Let SetLocalXML(xml)
	
		strLocalLandscapePATH = xml
		
	End Property 
	
	Public Property Get GetLocalXML
	
		GetLocalXML = strLocalLandscapePATH
		
	End Property 
	
	Public Property Let SetSystemName(sys)
	
		strSSN = UCase(sys)
		
	End Property 
	
	Public Property Let SetClientName(cli)
	
		strSCN = cli
	
	End Property 
	
	Public Property Get SessionFound
	
		SessionFound = boolSessionFound
		
	End Property 
	
		
	
End Class



Class RateUpload_v1

	Private oFSO
	Private oWSH
	Private oNET
	Private oSES ' Session should be obtained from SapLauncher.GetSession method
	Private strUserName ' System user name e.g. a293793
	Private strComputerName ' System name e.g. SKSENEW128
	
	' ============== Constructor & Destructor ==================
	Private Sub Class_Initialize
	
		Set oFSO = CreateObject("scripting.filesystemobject")
		Set oWSH = CreateObject("wscript.shell")
		Set oNET = CreateObject("wscript.network")
		oSES = Null 
		strUserName = oNET.UserName
		strComputerName = oNET.ComputerName

	End Sub
	
	Private Sub Class_Terminate
	
	End Sub
	
	' ============ P U B L I C  &  P R I V A T E   M E T H O D S  &   S U B R O U T I N E S =========



	' --------- UploadRates
	Public Function UploadRates(strFiles,strExRateType) ' strFiles is comma delimited list of files to upload
	
		Dim validfrom,SAPfile,i,ratetype,filename
		i = 0
		ratetype = UCase(strExRateType)
	
		For Each SAPfile In Split(strFiles,",")
		
			If oFSO.FileExists(SAPfile) Then 
				filename = oFSO.GetFileName(SAPfile) ' Returns 20200630.txt 
				validfrom = "" ' Clear
				validfrom = Mid(filename,7,2) & "." & Mid(filename,5,2) & "." & Mid(filename,1,4) ' SAP compatible date format DD.MM.YYYY
				KillPopups(oSES)
				oSES.findById("wnd[0]/tbar[0]/okcd").text = "/NZTC_ZCURR_UPLOAD"
				oSES.findById("wnd[0]").sendVKey 0 ' ENTER
				KillPopups(oSES)
				oSES.findById("wnd[0]/usr/txtP_FILE").text = SAPfile
				oSES.findById("wnd[0]/usr/txtP_KURST").text = ratetype
				oSES.findById("wnd[0]/usr/ctxtP_GDATU").text = validfrom
				oSES.findById("wnd[0]").sendVKey 8
				KillPopups(oSES)
				oSES.findById("wnd[0]").sendVKey 0
				KillPopups(oSES)
		
				Do While oSES.Children.Count > 1
					oSES.findById("wnd[0]").sendVKey 0
				Loop
				i = i + 1
				WScript.Sleep 2000 ' Wait a bit
			End If 	
		Next
		
		oSES.findById("wnd[0]/tbar[0]/okcd").text = "/NEX" ' Close transaction window
		oSES.findById("wnd[0]").sendVKey 0
	
		UploadRates = i ' Return the number of uploaded files or 0 if error occured   

	End Function 
	

				
''============================================================
'' Program:   SUB Killpopups
'' Desc:      Kill of SAP popup screens which could appear when executing SAP transactions
'' Called by: 
'' Call:      KillPopups
'' Arguments: s = connection.children(0)
'' Changes---------------------------------------------------
'' Date		Programmer	Change
'' 2020-06-01	Tomas Chudik(tomas.chudik@volvo.com)	Written as vbscript SUB with arguments; supports kill of "System Message", "Copyright", "License Information for Multiple Logon"
'' 2020-06-03	Tomas Chudik(tomas.chudik@volvo.com)	Version 2 without arguments for use in VBA
'' 2020-07-06	Tomas Chudik(tomas.chudik@volvo.com)	Version 3 (v3-arg) supports kill of Information window while ExRates rate adjustment
'' 2022-07-28	Tomas Chudik(tomas.chudik@volvo.com)	Correction in Multiple Logon condition
''============================================================

Public Function KillPopups(ByRef objSession)
		Do While objSession.Children.Count > 1
			If InStr(objSession.ActiveWindow.Text, "System Message") > 0 Then
				objSession.ActiveWindow.sendVKey 12
			ElseIf InStr(objSession.ActiveWindow.Text, "Information") > 0 And InStr(objSession.ActiveWindow.PopupDialogText, "Exchange rate adjusted to system settings") > 0 Then
				objSession.ActiveWindow.sendVKey 0
			ElseIf InStr(objSession.ActiveWindow.Text, "Copyright") > 0 Then
				objSession.ActiveWindow.sendVKey 0
			ElseIf InStr(objSession.ActiveWindow.Text, "License Information for Multiple Logon") > 0 Then
				objSession.ActiveWindow.findById("usr/radMULTI_LOGON_OPT2").select
				objSession.ActiveWindow.sendVKey 0
			'ElseIF   'Insert next type of popup windows which you want to kill
			Else
				Exit Do
			End If
		Loop
	End Function 
	' ================= P R O P E R T I E S ====================
	Public Property Let SAPSession(s)
		Set oSES = s
	End Property 	
		

	
End Class 





