'Replace $YOUR_VALUE_HERE token with actual values
Option Explicit
' Script error codes:
' 10x -> Login information problem
' 20x -> Command line parameters problem
' 40x -> SAP REST API errors
Const RESOURCE_NAME = "SPRESTAPI"
Const PROJECT = "RO01_SAPExRate"
Const DATE_RX = "^[0-9]{4}\.[0-9]{2}\.[0-9]{2}$"
Const WSCRIPT_RX = "wscript"
Const FOLDER = "C:\!AUTO\RO01_EXRATE"
Const PARAM_ONDEMAND_DATE = 1
Const PARAM_SAP_SYSTEM = 2
Const PARAM_SAP_CLIENT = 4
Const PARAMS_MANDATORY = 6
Const ForAppending = 8
Const ForReading = 1
Const ForWriting = 2
'############################################
'############ V a r i a b l e s #############
'############################################
Dim actualParameters : actualParameters = 0 ' Bitfield like
Dim paramSapSystem,paramSapClient
Dim oSPL : Set oSPL = New SAPLauncher
Dim strDate
Dim isGui : isGui = False 
Dim oTargetDate
Dim oFirstNonTcdDate
Dim tcdString
Dim oSES : oSES = Null 
Dim oWSH : Set oWSH = CreateObject("Wscript.Shell")
Dim oFSO : Set oFSO = CreateObject("Scripting.FileSystemObject")
Dim oRX : Set oRX = New RegExp
Dim oSL : Set oSL = New SAPLauncher
Dim oDF : Set oDF = New DateFormatter
Dim oTCD : Set oTCD = New TCDCalendar
Dim oMAIL : Set oMAIL = New Mailer
Dim oRATECHECK : Set oRATECHECK = New RateCheck_v1
Dim oRATES
Dim oFILE
oMAIL.AddAdmin = $YOUR_VALUE_HERE
oRX.Pattern = DATE_RX
oRX.Global = True
oRX.IgnoreCase = True



'Date for which exrate processing should be performed. 
'A) If script is executed w/o CLI params it is Today
'B) If script is executed w/ CLI params we look for -o DD.MM.YYYY 
'C) If script is executed interactively (wscript) users inputs values via a inputbox
'By default it is set to Today 
oTargetDate = CDate(Right("0000" & Year(Date),4) & "-" & Right("00" & Month(Date),2) & "-" & Right("00" & Day(Date),2) + 1) ' Normal automated workflow. Target date is tomorrow.
'######################################
'############# M A I N ################
'######################################
'++++++++++++++++++++++++++++++
'+ CLI or GUI
'++++++++++++++++++++++++++++++
oRX.Pattern = WSCRIPT_RX
If oRX.Test(WScript.FullName) Then	
	isGui = True
End If
oRX.Pattern = DATE_RX
'++++++++++++++++++++++++++++
'+ Process CLI params
'++++++++++++++++++++++++++++
If Not isGui Then ' Process CLI parameters
	Dim arg,args
	args = WScript.Arguments.Count
	For arg = 0 To WScript.Arguments.Count - 1
	
		Select Case LCase(WScript.Arguments.Item(arg))
		
			Case "-o","--ondemand"
			
				If arg + 1 < args Then
					
					If Not oRX.Test(WScript.Arguments.Item(arg + 1)) Then
						WScript.Echo "Invalid date format. Accepted format is: YYYY.MM.DD"
						WScript.Quit(202) ' Invalid date format	
					Else
						oTargetDate = CDate(Replace(WScript.Arguments.Item(arg + 1),".","-")) ' Overwrite the default value - Today
						actualParameters = actualParameters Or PARAM_ONDEMAND_DATE
					End If 
					
				Else
					WScript.Echo "Missing parameter value '-o|--ondemand YYYY.MM.DD'"
					WScript.Quit(201)
				End If 
				
			Case "-s","--system"
			
				If arg + 1 < args Then 
				
					paramSapSystem = WScript.Arguments.Item(arg + 1)
					actualParameters = actualParameters Or PARAM_SAP_SYSTEM
					
				End If 
				
			Case "-c","--client"
			
				If arg + 1 < args Then
				
					paramSapClient = WScript.Arguments.Item(arg + 1)
					actualParameters = actualParameters Or PARAM_SAP_CLIENT
					
				End If 
					
		End Select 
						
	Next
	
Else ' Prompt user for input via GUI
	PromptGuiParameters
End If
'+++++++++++++++++++++++++++++++++++++++++++++++++
'+ Verify mandatory parameters have been provided
'+++++++++++++++++++++++++++++++++++++++++++++++++
If Not isGui Then 
	If Not actualParameters And PARAMS_MANDATORY Then
		WScript.Echo "Missing mandatory parameters"
		WScript.Echo "Usage: " & WScript.ScriptName & " -s|--system $SAP_SYSTEM_NAME -c|--client $SAP_CLIENT_NAME [-o|--ondemand $DATE_IN_YYYY.MM.DD]"
		WScript.Echo "Example 1: " & WScript.ScriptName & " -s FQ2 -c 105 -o 2022.01.05"
		WScript.Quit(200) ' Missing mandatory parameters
	End If 
	
	WScript.Echo "All mandatory parameters provided: " & CBool(actualParameters And PARAMS_MANDATORY)
	WScript.Echo "Ondemand requested: " & CBool(actualParameters And PARAM_ONDEMAND_DATE)
	WScript.Echo "SAP System: " & paramSapSystem
	WScript.Echo "SAP Client: " & paramSapClient
End If 
'++++++++++++++++++++++++++++
'+ Create the output folder +
'++++++++++++++++++++++++++++
If Not oFSO.FolderExists(oFSO.GetParentFolderName(FOLDER)) Then
	oFSO.CreateFolder(oFSO.GetParentFolderName(FOLDER))
End If 
If Not oFSO.FolderExists(FOLDER) Then
	oFSO.CreateFolder(FOLDER)
End If
'++++++++++++++++++++++++++++++++
'+ Determine the correct dates  +
'++++++++++++++++++++++++++++++++
'Intialize TCD calendar
oTCD.Init "Provider=Microsoft.ACE.OLEDB.12.0;WSS;IMEX=1;RetrieveIds=Yes;" & _
		  "DATABASE=https://$YOUR_VALUE_HERE.sharepoint.com/sites/unit-financean/exrate/RO01_exrate;" & _
		  "LIST=RO01_exrate_calendar;"
		  

oFirstNonTcdDate = oTCD.firstNTCD(oTargetDate - 1) 'Subtract 1 so that we don't need to check if this is ondemand or not. If target date is Target=Today + 1 (normal automated workflow when today we upload rates for tomorrow) 
												   'then rates are published at least one day prior. The very same thing applies if target is specified in ondemand parameter. 
												   'We always have to go AT LEAST one day back to find the date the rates had been published.
												   'Given ondemand parameter value 2023.01.09 we're going to be looking for the date rates were published for 2023.01.09. We go one day back since rates are always published one day
												   'prior. If this day was a weekend (Saturday or Sunday) OR a holiday, we go one day back and repeat the process until we find such a date, that was neither the weekend nor holiday


If Not isGui Then
    WScript.Echo "Target date (non formatted): " & oTargetDate
	WScript.Echo "Target date (formatted): " & oDF.ToDayMonthYearWithDots(oTargetDate)
	WScript.Echo "First Non TCD date (non formatted): " & oFirstNonTcdDate
	WScript.Echo "First Non TCD date (formatted): " & oDF.ToDayMonthYearWithDots(oFirstNonTcdDate)
End If 
'+++++++++++++++++++++++++++++
'+  Download exchange rates  +
'+++++++++++++++++++++++++++++
Dim oHTTP : Set oHTTP = CreateObject("MSXML2.XMLHTTP")
Dim oXML : Set oXML = CreateObject("MSXML2.DOMDOCUMENT")

With oHTTP
	.open "GET","https://bnr.ro/files/xml/years/nbrfxrates" & Year(oFirstNonTcdDate) & ".xml",False ' Open XML for the year where our non TCD date exits. This is needed at the end of the year
	.setRequestHeader "accept","application/xml"
	.send
End With 

If Not oHTTP.status = 200 Then
	If Not isGui Then
		WScript.Echo "HTTP status: " & oHTTP.status
		WScript.Echo "HTTP errorMessage: " & oHTTP.responseText
	End If
	
	WScript.Quit oHTTP.status
	
End If


oXML.loadXML oHTTP.responseText
oXML.setProperty "SelectionNamespaces", "xmlns=""http://www.bnr.ro/xsd"" xmlns:xsi=""http://www.w3.org/2001/XMLSchema-instance"""

Dim oCube : Set oCube = oXML.selectSingleNode("//Cube[@date=""" & oDF.ToYearMonthDayWithDashes(oFirstNonTcdDate) & """]")
On Error Resume Next 

'If date is not found within the xml this call should generate excepction we "catch" 
Dim childrenLen : childrenLen = oCube.childNodes.length

If err.number <> 0 Then
 	If Not isGui Then
		WScript.Echo "Rates not yet published"
		WScript.Quit
	End If 
	
	MsgBox "Rates not yet available",vbOKOnly + vbInformation,"Rates not yet published"
	WScript.Quit
End If 

'Date has been found within the xml. Continue
Set oFILE = oFSO.OpenTextFile(FOLDER & "\" & oDF.ToYearMonthDay(oTargetDate) & ".txt",ForWriting,True)
oFILE.WriteLine "EUR	RON	" & FormatNumber(Replace(oXML.selectSingleNode("//Cube[@date=""" & oDF.ToYearMonthDayWithDashes(oFirstNonTcdDate) & """]/Rate[@currency=""EUR""]").text,".",","),5) & "	1	1"
oFILE.WriteLine "HUF	RON	" & FormatNumber(Replace(oXML.selectSingleNode("//Cube[@date=""" & oDF.ToYearMonthDayWithDashes(oFirstNonTcdDate) & """]/Rate[@currency=""HUF""]").text,".",","),5) & "	1	1"
oFILE.WriteLine "SEK	RON	" & FormatNumber(Replace(oXML.selectSingleNode("//Cube[@date=""" & oDF.ToYearMonthDayWithDashes(oFirstNonTcdDate) & """]/Rate[@currency=""SEK""]").text,".",","),5) & "	1	1"
oFILE.WriteLine "USD	RON	" & FormatNumber(Replace(oXML.selectSingleNode("//Cube[@date=""" & oDF.ToYearMonthDayWithDashes(oFirstNonTcdDate) & """]/Rate[@currency=""USD""]").text,".",","),5) & "	1	1"
oFILE.Close
'+++++++++++++++++++++++++++
'+  Upload exchange rates  +
'+++++++++++++++++++++++++++
oSPL.SetClientName = paramSapClient
oSPL.SetSystemName = paramSapSystem
oSPL.SetLocalXML = oWSH.ExpandEnvironmentStrings("%APPDATA%") & "\SAP\Common\SAPUILandscape.xml"
oSPL.CheckSAPLogon
oSPL.FindSAPSession

If Not oSPL.SessionFound Then 
	debug.WriteLine "Session not found"
	WScript.Quit(1)
'	SEND message
'	LOG
End If 

Set oSES = oSPL.GetSession
oSPL.KillPopups(SS)

oSES.StartTransaction "YTC_EXRATE_RO01"
oSES.findById("wnd[0]/usr/ctxtP_FILE").text = FOLDER & "\" & oDF.ToYearMonthDay(oTargetDate) & ".txt"
oSES.findById("wnd[0]/usr/txtP_KURST").text = "YRO1"
oSES.findById("wnd[0]/usr/ctxtP_GDATU").text = oDF.ToDayMonthYearWithDots(oTargetDate)
oSES.findById("wnd[0]").sendVKey 8
oSES.findById("wnd[0]").sendVKey 0
'++++++++++++++++++++++++++++++
'+ Check rates in SAP         +
'++++++++++++++++++++++++++++++
oRATECHECK.SAPSession = oSES
oRATECHECK.Init "RON" ' Home currency
oRATECHECK.CheckRates FOLDER  & "\" & oDF.ToYearMonthDay(oTargetDate) & ".txt", "YRO1"

If oRATECHECK.FilesVerified = 0 Then 
	oMAIL.SendMessage file & " not found","E",oSPL.SAPsysName
	WScript.Quit
End If 

If oRATECHECK.GetNumIncompleteEntries > 0 Or oRATECHECK.GetNumInvalidEntries > 0 Or oRATECHECK.GetNumMissingEntries > 0 Then
	oMAIL.SendMessage "Verification failed in " & oSPL.SAPsysName & ". MISSING FILES: " & oRATECHECK.GetMissingEntries & " INCOMPLETE FILES: " & oRATECHECK.GetIncompleteEntries & " INVALID FILES: " & oRATECHECK.GetInvalidEntries,"E",oSPL.SAPsysName
	WScript.Quit
Else
	oMAIL.SendMessage FOLDER & "\" & oDF.ToYearMonthDay(oTargetDate) & ".txt" & " successfully verified","I",oSPL.SAPsysName
End If 

Checkin PROJECT, RESOURCE_NAME
WScript.Quit 0 ' OK 




'################################################
'###### C L A S S   D E F I N I T I O N S #######
'################################################

Sub PromptGuiParameters

	Dim answerDate : answerDate = InputBox("Insert date in the format: YYYY.MM.DD","Select date",Year(Date) & "." & Right("00" & Month(Date),2) & "." & Right("00" & Day(Date),2))
	
	If answerDate = "" Then
		WScript.Quit
	End If 
	
	Do While Not oRX.Test(answerDate)
		MsgBox "Invalid date format",vbOKOnly + vbCritical,"Invalid date format"
		answerDate = InputBox("Insert date in the format: YYYY.MM.DD","Select date",Year(Date) & "." & Right("00" & Month(Date),2) & "." & Right("00" & Day(Date),2))
		
		If answerDate = "" Then
			WScript.Quit
		End If 
	Loop
	
	oTargetDate = CDate(Replace(answerDate,".","-"))
	
	paramSapSystem = InputBox("Insert SAP system name, e.g fq2","Select SAP system","FQ2")
	If paramSapSystem = "" Then
		WScript.Quit
	End If 
	actualParameters = actualParameters Or PARAM_SAP_SYSTEM
	
	paramSapClient = InputBox("Insert SAP client name, e.g 105","Select SAP client","105")
	If paramSapClient = "" Then
		WScript.Quit
	End If
	actualParameters = actualParameters Or PARAM_SAP_CLIENT
	
	If MsgBox("Are these values correct ?" & vbCrLf & vbCrLf & "SAP System: " & paramSapSystem _
		& vbCrLf & "SAP client: " & paramSapClient & vbCrLf & "Date: " _
		& Year(oTargetDate) _
		& "." & Right("00" & Month(oTargetDate),2) _
		& "." & Right("00" & Day(oTargetDate),2),vbYesNo + vbQuestion,"Input verification") = vbNo Then
	
		If MsgBox("",vbRetryCancel,"") = vbRetry Then
			PromptGuiParameters
		Else
			WScript.Quit
		End If
	Else
		Exit sub
	End If 
 	
End Sub 


Class TCDCalendar
	'v3
	'Utilizes a sharepoint list as s source of TCDs
	'Private member variables 
	Private oRX__ 
	Private oCON__
	Private oRST__
	Private oDICT__ ' Dictionary holding closing days key=day value=name of holiday
	Private strConnectionString__
	Private strList__
	Private daySunday__
	Private daySaturday__
	Private oFirstNTCD__
	
	
	Private Sub Class_Initialize
		
		Set oRX__ = New RegExp
		oRX__.Pattern = "(?:list)=\w*"
		oRX__.Multiline = False
		oRX__.IgnoreCase = True
		oRX__.Global = True
		Set oDICT__ = CreateObject("scripting.dictionary")
		Set oRST__ = CreateObject("adodb.recordset")
		Set oCON__ = CreateObject("adodb.connection")
		strConnectionString__ = ""
		strList__ = ""
		daySaturday__ = 7
		daySunday__ = 1
		
	End Sub 
	
    
	Public Function Init(sConnectionString)
	
	    Dim matches : Set matches = oRX__.Execute(sConnectionString)
	    oRX__.Pattern = "list="
	    strList__ = oRX__.Replace(matches.Item(0),"")
		strConnectionString__ = sConnectionString
		oCON__.ConnectionString = strConnectionString__
		
	
		If Not strList__ = "" And Not IsNull(strConnectionString__) And Not strConnectionString__ = "" Then
		   
		    oCON__.Open
	    	oRST__.Open "SELECT Title, TCD FROM [" & strList__ & "]", oCON__, 3, 3
	        oRST__.MoveFirst
	        
	        Do While Not oRST__.EOF
	        
	        	oDICT__.Add oRST__.Fields("TCD").Value, oRST__.Fields("Title").Value
	        	oRST__.MoveNext
	        	
	        Loop
	        
	    Else
	    
	    	Init = -1
	    	
	    End If
	    	    	
    End Function 
    
    
    Private Function FindFirstNonTcdDate(D)
    
    	oFirstNTCD__ = D
    	    
        If (Weekday(D) = daySaturday__ Or Weekday(D) = daySunday__) Or oDICT__.Exists(D) Then
        	FindFirstNonTcdDate(D - 1) ' Recursive call
        Else
        	oFirstNTCD__ = D
        	Exit Function
        End If 
        
    End Function 
	
	Public Property Get connectionString
	
		connectionString = strConnectionString__
		
	End Property 
	
	Public Property Get list
	   
	    list = strList__
  		
	End Property
	
	Public Property Get firstNTCD(D)
	
		FindFirstNonTcdDate(D)
	
		firstNTCD = oFirstNTCD__
		
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
	
	Public Function ToYearMonthDayHourMinuteSecondWithZeros(D,T)
		ToYearMonthDayHourMinuteSecondWithZeros = Right("0000" & Year(D),4) & Right("00" & Month(D),2) & Right("00" & Day(D),2) & Right("00" & Hour(T),2) & Right("00" & Minute(T),2) & Right("00" & Second(T),2)
	End Function
	
	
End Class 





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
		
		Set oSAPGUI = GetObject("SAPGUI") ' This fails is saplogon is not running. We're connecting to the COM object not creating our own instance in this approach
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
	Public Function UploadRates(strFiles,strExRateType,boolDoNotNEX) ' strFiles is comma delimited list of files to upload,ex rate type ie YHR2, preserve session. Do not cal /NEX
	
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
		
		If Not boolDoNotNEX Then
			oSES.findById("wnd[0]/tbar[0]/okcd").text = "/NEX" ' Close transaction window
			oSES.findById("wnd[0]").sendVKey 0
		End If 
	
		UploadRates = i ' Return the number of uploaded files or 0 if error occured   

	End Function 
	
	
End Class 




Class RateCheck_v1
	
	
	Private oFSO
	Private oWSH
	Private oSES
	Private oFile
	Private oTempFile        ' temp file to hold data from sap TCURR
	Private strHomeCurrency  ' Home currency e.g. CZK
	Private numFilesVerified ' Number of successfully verified files
	Private strTempFilePath  ' Absolute path to the temp file
	Private strTempFileName  ' temp file name
	Private strIncompleteEntries ' Input line count and sap entries count don't match
	Private strMissingEntries ' There is no rate for this day
	Private strInvalidEntries ' These entries line count match, but lines differ i.e rates are invalid
	Private missingEntries 
	Private incompleteEntries
	Private invalidEntries
	
	
	' ============== Constructor & Destructor ==================
	Private Sub Class_Initialize
	
		Set oFSO = CreateObject("scripting.filesystemobject")
		Set oWSH = CreateObject("wscript.shell")
		numFilesVerified = 0
		missingEntries = 0
		incompleteEntries = 0
		invalidEntries = 0
		oSES = Null 
		strMissingEntries = Null
		strIncompleteEntries = Null
		strInvalidEntries = Null
		strTempFilePath = Null
		strTempFileName = Null
		strHomeCurrency = Null

	End Sub
	
	Private Sub Class_Terminate
	
	End Sub
	
	' ============ P U B L I C  &  P R I V A T E   M E T H O D S  &   S U B R O U T I N E S ===========
	
	Public Function Init(str_home_curr)
	
		strHomeCurrency = str_home_curr
		
	End Function 
	
	' ------------ CreateGUID
	
	Private Function CreateGUID
  		Dim TypeLib
  		Set TypeLib = CreateObject("Scriptlet.TypeLib")
  		CreateGUID = Mid(TypeLib.Guid, 2, 36)
	End Function



	' --------- CheckRates
	Public Function CheckRates(strFiles,strExRateType) ' strFiles is comma delimited list of files to check
	
		Dim SAPfile,files
		
		files = Split(strFiles,",") ' Split files and use the first one to determine where to put temp file
		
		
		If strHomeCurrency = Null Or strHomeCurrency = "" Then
			CheckRates = -1 ' ERROR, home currency not set
		End If 
		
		strTempFileName = CreateGUID ' Create a temp file name
		
		strTempFilePath = oFSO.GetParentFolderName(files(0)) ' Get a temp file location
		
	
		For Each SAPfile In Split(strFiles,",")
			
			If oFSO.FileExists(SAPfile) Then 
				Check SAPfile,strExRateType
			End If 
			
		Next
		
		 
		oSES.findById("wnd[0]/tbar[0]/okcd").text = "/NEX" ' Close transaction window
		oSES.findById("wnd[0]").sendVKey 0
		
		If oFSO.FileExists(strTempFilePath & "\" & strTempFileName & ".txt") Then 
		
			oFSO.DeleteFile strTempFilePath & "\" & strTempFileName & ".txt"
			
		End If 
		CheckRates = numFilesVerified ' Returns number of successfully verified files.
		

	End Function 
	
	
	
	
	
	
	
	
	Private Sub Check(strFile,strType) ' Private sub to check files. Call within for loop
	
		Dim lines,filename,gdatu,line,i,sapline,fileline,column,sapentries,j,saplinetrimmed,filelinetrimmed
		i = 0
		lines = 0
		
		If Not oFSO.FileExists(strFile) Then
			Exit Sub 
		End If 
		 
		 Set oFile = oFSO.OpenTextFile(strFile,1,False) ' Open file containing uploaded rates for reading
		 
		 Do While Not oFile.AtEndOfStream
		 	oFile.ReadLine
		 Loop
		 	
		 lines = oFile.Line - 1
		 oFile.Close
		 	
		 	
		filename = oFSO.GetBaseName(strFile) ' Returns 20200630 
		gdatu = 99999999 - filename
		KillPopups(oSES)
		oSES.findById("wnd[0]/tbar[0]/okcd").text = "/nse17"
		oSES.findById("wnd[0]").sendVKey 0 ' ENTER
		KillPopups(oSES)
		oSES.findById("wnd[0]/usr/ctxtDD02V-TABNAME").text = "TCURR"
		oSES.findById("wnd[0]").sendVKey 0 ' ENTER
		KillPopups(oSES)
		' FIELDS
		oSES.findById("wnd[0]/usr/tblSAPMSTAZTABCTRL200/ctxtRSTAZ-FSELECT[1,1]").text = LCase(strType)
		oSES.findById("wnd[0]/usr/tblSAPMSTAZTABCTRL200/ctxtRSTAZ-FSELECT[1,2]").text = ""
		oSES.findById("wnd[0]/usr/tblSAPMSTAZTABCTRL200/ctxtRSTAZ-FSELECT[1,3]").text = ""
		oSES.findById("wnd[0]/usr/tblSAPMSTAZTABCTRL200/ctxtRSTAZ-FSELECT[1,4]").text = gdatu
		' FLAGS
		oSES.findById("wnd[0]/usr/tblSAPMSTAZTABCTRL200/ctxtRSTAZ-SHOWFLAG[2,0]").text = ""  ' CLIENT
		oSES.findById("wnd[0]/usr/tblSAPMSTAZTABCTRL200/ctxtRSTAZ-SHOWFLAG[2,1]").text = ""  ' KURST
		oSES.findById("wnd[0]/usr/tblSAPMSTAZTABCTRL200/ctxtRSTAZ-SHOWFLAG[2,2]").text = "X" ' FCURR
		oSES.findById("wnd[0]/usr/tblSAPMSTAZTABCTRL200/ctxtRSTAZ-SHOWFLAG[2,3]").text = "X" ' TCURR
		oSES.findById("wnd[0]/usr/tblSAPMSTAZTABCTRL200/ctxtRSTAZ-SHOWFLAG[2,4]").text = ""  ' GDATU
		oSES.findById("wnd[0]/usr/tblSAPMSTAZTABCTRL200/ctxtRSTAZ-SHOWFLAG[2,5]").text = "X" ' UKURS
		oSES.findById("wnd[0]").sendVKey 8
		KillPopups(oSES)
		
		If oSES.findById("wnd[0]/sbar/pane[0]").text <> "" Or oSES.findById("wnd[0]/sbar/pane[0]").text = "No values selected in the specified area" Then
			
			missingEntries = missingEntries + 1 
			strMissingEntries = strMissingEntries & " " & strFile
			numFilesVerified = numFilesVerified + 1
			Exit Sub 
		
		End If 
			
		sapentries = oSES.findById("wnd[0]/usr/lbl[23,3]").text ' Number of entries. Compare this to the input file line count
		
		If CInt(sapentries) <> lines Then
		  
			incompleteEntries = incompleteEntries + 1 
			strIncompleteEntries = strIncompleteEntries & " " & strFile
			numFilesVerified = numFilesVerified + 1
			Exit Sub 
			
		End If 
		
		' Continue with complete entries. Generate output file from SAP
		oSES.findById("wnd[0]/mbar/menu[5]/menu[5]/menu[2]/menu[1]").select
		oSES.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[1,0]").select
		oSES.findById("wnd[1]").sendVKey 0
		KillPopups(oSES)
		oSES.findById("wnd[1]/usr/ctxtDY_PATH").text = strTempFilePath 'directory
		oSES.findById("wnd[1]/usr/ctxtDY_FILENAME").text = strTempFileName & ".txt"
		oSES.findById("wnd[1]").sendVKey 11
		KillPopups(oSES)
	
		Set oFile = oFSO.OpenTextFile(strFile,1,False) ' Open input file for reading
		Set oTempFile = oFSO.OpenTextFile(strTempFilePath & "\" & strTempFileName & ".txt") ' Open sap generated file for reading
		j = 0
		Do While j < 9
			oTempFile.SkipLine
			j = j + 1
		Loop 
		
		Do While Not oFile.AtEndOfStream
		
			
			sapline = Split(oTempFile.ReadLine,vbCrLf)
			fileline = Split(oFile.ReadLine,vbCrLf)
			column = Split(fileline(0),vbTab)
			saplinetrimmed = Replace((Trim(sapline(0))),vbTab,"")
			filelinetrimmed = Replace((Trim(column(0) & column(1) & column(2))),vbTab,"")
			
			If Replace(saplinetrimmed," ","") <> filelinetrimmed Then
			
				invalidEntries = invalidEntries + 1
				numFilesVerified = numFilesVerified + 1
				strInvalidEntries = strInvalidEntries & " " & strFile
				oFile.Close
				oTempFile.Close
				Exit Sub 
		
			End If 
			
			
		Loop
		
		
		numFilesVerified = numFilesVerified + 1
		oTempFile.Close ' Close the temp file
		oFile.Close     ' Close the rate file
		
	End Sub  	
	
	
				
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
	
	Public Property Get FilesVerified
		FilesVerified = numFilesVerified
	End Property 
	
	Public Property Get GetMissingEntries
		GetMissingEntries = strMissingEntries
	End Property  
	
	Public Property Get GetIncompleteEntries
		GetIncompleteEntries = strIncompleteEntries
	End Property 
	
	Public Property Get GetNumMissingEntries
		GetNumMissingEntries = missingEntries
	End Property
	
	Public Property Get GetNumIncompleteEntries
		GetNumIncompleteEntries = incompleteEntries
	End Property 
	
	Public Property Get GetNumInvalidEntries
		GetNumInvalidEntries = invalidEntries
	End Property 
	
	Public Property Get GetInvalidEntries
		GetInvalidEntries = strInvalidEntries
	End Property 
	
End Class 


Class Mailer

	Private oEmail
	Private oSysInfo
	Private oUser
	Private strAdmins
	Private oNET
	Private strUserName
	Private strComputerName
	
	Private Sub Class_Initialize
	
		Set oEmail = CreateObject("CDO.Message")
		Set oSysInfo = CreateObject("ADSystemInfo")
		Set oUser = GetObject("LDAP://" & oSysInfo.UserName)
		Set oNET = CreateObject("Wscript.Network")
		strUserName = oNET.UserName
		strComputerName = oNET.ComputerName
		
	End Sub 
	
	Private Sub Class_Terminate
	
	End Sub 
	
	
	Public Function SendMessage(strMessage,strSeverity,strAppendToSubject)
	
		Dim admin,strFrom,oUser
		Set oUser = GetObject("LDAP://" & oSysInfo.UserName)
		For Each admin In Split(GetAdmins,",")
			
			With oEmail 
				.From = oUser.Mail
				.To = admin
				.Subject = strSeverity & ";" & WScript.ScriptName & ";" & Year(Date()) & "-" & right("00" & Month(Date()),2) & "-" & right("00" & Day(Date()),2) & ";" & Time() & ";" & strUserName & ";" & strComputerName & ";" & strAppendToSubject
				.Configuration.Fields.Item _
   				("http://schemas.microsoft.com/cdo/configuration/sendusing") = 2
				.Configuration.Fields.Item _
    			("http://schemas.microsoft.com/cdo/configuration/smtpserver") = _
        		"$YOUR_VALUE_HERE" 
				.Configuration.Fields.Item _
  	    		("http://schemas.microsoft.com/cdo/configuration/smtpserverport") = 25
  	    		.TextBody = strMessage
  	    		.Configuration.Fields.Item("urn:schemas:mailheader:X-MSMail-Priority") = "High"
  	    		.Configuration.Fields.Item("urn:schemas:httpmail:importance") = 2
				.Configuration.Fields.Item("urn:schemas:mailheader:X-Priority") = 2
				.Configuration.Fields.Update
				.Send
			End With 
			
		Next
	
	End Function
	
	Public Property Let AddAdmin(strEmailAddress)
		strAdmins = strAdmins & strEmailAddress & ","
	End Property 
	
	Public Property Get GetAdmins
		GetAdmins = Left(strAdmins,Len(strAdmins) - 1)
	End Property 
	
End Class 


'Function for checkig-in to the watchdog list
Function Checkin(sProjectName,sResourceName)
	Dim sUserName,sUserSecret,sSiteUrl,sDomain,sTenantID,sClientID,sXDigest,sAccessToken,tmp,rxResult
	Dim oHTTP : Set oHTTP = CreateObject("MSXML2.ServerXMLHTTP.3.0")
	Dim oXML : Set oXML = CreateObject("MSXML2.DOMDOCUMENT")
	Dim oCON : Set oCON = CreateObject("Adodb.Connection")
	Dim oRST : Set oRST = CreateObject("Adodb.Recordset")
	Dim oRX : Set oRX = New RegExp
	Dim connectionString : connectionString = "Provider=Microsoft.ACE.OLEDB.12.0;WSS;IMEX=1;RetrieveIds=Yes;" & _
							 "DATABASE=https://$YOUR_VALUE_HERE.sharepoint.com/sites/unit-rc-sk-bs-it/CREDENTIALS;" & _
						 	 "LIST=CREDENTIALS;"
	
	'Load credentials from SP list
	oCON.ConnectionString = connectionString
	oCON.Open
	oRST.Open "SELECT Host,Username,Password FROM [CREDENTIALS] WHERE Title='" & sResourceName & "';", oCON, 3, 3
	
	If oRST.EOF Or oRST.BOF Then
		Checkin = 0
		Exit Function 
	End If 
     
	oRX.Pattern = "^(?:https?:\/\/)?(?:[^@\n]+@)?(?:www\.)?([^:\/\n?]+)"
	oRX.Global = True
	
	oRST.MoveFirst
	
	sDomain = oRX.Execute(oRST.Fields("Host").Value)(0)
	oRX.Pattern = "(http:\/\/|https:\/\/)"
	
	sDomain = oRX.Replace(sDomain,"")
	sUserName = oRST.Fields("Username").Value
	sUserSecret = oRST.Fields("Password").Value
	sSiteUrl = oRST.Fields("Host").Value
	
	oRST.Close

	'Get TenantID & ClientID/ResourceID
	On Error Resume Next 
	With oHTTP
		.open "GET",sSiteUrl & "/_vti_bin/client.svc",False
		.setRequestHeader "Authorization","Bearer"
		.send
	End With
		
	oRX.Pattern = "Bearer realm=""([a-zA-Z0-9]{1,}-)*[a-zA-Z0-9]{12}"
	Set rxResult = oRX.Execute(oHTTP.getResponseHeader("WWW-Authenticate"))
	oRX.Pattern = "[a-fA-F0-9]{8}-([a-fA-F0-9]{4}-){3}[a-fA-F0-9]{12}"
	sTenantID = oRX.Execute(rxResult(0))(0)
	
	oRX.Pattern = "client_id=""[a-fA-F0-9]{8}-([a-fA-F0-9]{4}-){3}[a-fA-F0-9]{12}"
	Set rxResult = oRX.Execute(oHTTP.getResponseHeader("WWW-Authenticate"))
	oRX.Pattern = "[a-fA-F0-9]{8}-([a-fA-F0-9]{4}-){3}[a-fA-F0-9]{12}"
	sClientID = oRX.Execute(rxResult(0))(0)
	
	'Get AccessToken
	Dim sBody : sBody = "grant_type=client_credentials&client_id=" & sUserName & "@" & sTenantID & "&client_secret=" & sUserSecret & "&resource=" & sClientID & "/" & sDomain & "@" & sTenantID
	With oHTTP
		.open "POST", "https://accounts.accesscontrol.windows.net/" & sTenantID & "/tokens/OAuth/2", False
		.setRequestHeader "Host","accounts.accesscontrol.windows.net"
		.setRequestHeader "Content-Type","application/x-www-form-urlencoded"
		.setRequestHeader "Content-Length", CStr(Len(sBody))
		.send sBody
	End With 
	
	oRX.Pattern = "access_token"":"".*"
	Set rxResult = oRX.Execute(oHTTP.responseText)
	rxResult = Split(rxResult(0),":")
	rxResult(1) = Replace(rxResult(1),"""","")
	rxResult(1) = Replace(rxResult(1),"}","")
	sAccessToken = rxResult(1) ' Save the token 
	
	'Get XDigest
	With oHTTP
		oHTTP.open "POST", sSiteUrl & "/_api/contextinfo", False 
		oHTTP.setRequestHeader "accept","application/atom+xml;odata=verbose"
		oHTTP.setRequestHeader "authorization", "Bearer " & sAccessToken
		oHTTP.send
	End With 
	
	oXML.loadXML oHTTP.responseText
	sXDigest = oXML.selectSingleNode("//d:FormDigestValue").text
	
	'Send query
	With oHTTP
		.open "GET", sSiteUrl & "/_api/web/lists/getbytitle('WDAPP')/items?$select=Title&$filter=(Title eq '" & sProjectName & "')", False        
		.setRequestHeader "Authorization", "Bearer " & sAccessToken
		.setRequestHeader "Accept", "application/atom+xml;odata=verbose"
		.setRequestHeader "X-RequestDigest", sXDigest
		.send
	End With 
	
	'Patch record
	Dim oNet : Set oNet = CreateObject("WScript.Network")
	Dim oSysInfo : Set oSysInfo = CreateObject("ADSystemInfo")
	Dim oLDAP : Set oLDAP = GetObject("LDAP://" & oSysInfo.UserName)
	oXML.loadXML oHTTP.responseText
	Dim url : url = oXML.selectSingleNode("//feed").attributes.getNamedItem("xml:base").text
	url = url & oXML.selectSingleNode("//entry/link[@rel=""edit""]").attributes.getNamedItem("href").text
  	
  
	With oHTTP
		.open "PATCH", url, False
		.setRequestHeader "Accept","application/json;odata=verbose"
		.setRequestHeader "Content-Type","application/json"
		.setRequestHeader "Authorization","Bearer " & sAccessToken
		.setRequestHeader "If-Match","*"
		.send "{""ComputerName"":""" & Trim(oNet.ComputerName) & """,""UserName"":""" & Trim(oLDAP.displayName) & """,""UserID"":""" & Trim(oLDAP.sAMAccountName) & """}"
	End With
	

	
End Function
