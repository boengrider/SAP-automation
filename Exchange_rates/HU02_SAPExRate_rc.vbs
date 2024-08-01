Option Explicit

Const RESOURCE_NAME = "SPRESTAPI"
Const PROJECT = "HU02_SAPExRate"
Const CRFILE = "C:\!AUTO\CREDENTIALS\logins.txt"
Const DATE_RX = "^[1-9]{1}[0-9]{3}\.([0]{1}[1-9]{1}|[1]{1}[0-2]{1})\.([0]{1}[1-9]{1}|[1-2]{1}[0-9]{1}|[3]{1}[0-1]{1})$"
Const FOLDER = "C:\ExRate\HU02"
Const ForAppending = 8
Const ForReading = 1
Const ForWriting = 2
'############################################
'############ V a r i a b l e s #############
'############################################

Dim HU_TCD : HU_TCD = "01012022;02012022;08012022;09012022;15012022;16012022;22012022;23012022;29012022;" _
             & "30012022;05022022;06022022;12022022;13022022;19022022;20022022;26022022;27022022;" _
             & "05032022;06032022;12032022;13032022;14032022;15032022;19032022;20032022;27032022;" _
             & "02042022;03042022;09042022;10042022;15042022;16042022;17042022;18042022;23042022;" _
             & "24042022;30042022;01052022;07052022;08052022;14052022;15052022;21052022;22052022;" _
             & "28052022;29052022;04062022;05062022;06062022;11062022;12062022;18062022;19062022;" _
             & "25062022;26062022;02072022;03072022;09072022;10072022;16072022;17072022;23072022;" _
             & "24072022;30072022;31072022;06082022;07082022;13082022;14082022;20082022;21082022;" _
             & "27082022;28082022;03092022;04092022;10092022;11092022;17092022;18092022;24092022;" _
             & "25092022;01102022;02102022;08102022;09102022;16102022;22102022;23102022;29102022;" _
             & "30102022;31102022;01112022;05112022;06112022;12112022;13112022;19112022;20112022;" _
             & "26112022;27112022;03122022;04122022;10122022;11122022;17122022;18122022;24122022;" _
             & "25122022;26122022;31122022;01012023;020122023"
Dim strDate
Dim oTargetDate
Dim tcd
Dim oSES : oSES = Null 
Dim oWSH : Set oWSH = CreateObject("Wscript.Shell")
Dim oFSO : Set oFSO = CreateObject("Scripting.FileSystemObject")
Dim oRX : Set oRX = New RegExp
Dim oSAPLAUNCHER
Dim oRATECHECK
Dim oMAIL : Set oMAIL = New Mailer
Dim oDF : Set oDF = New DateFormatter
Dim oTCD : Set oTCD = New TCDCalendar
Dim oRATES : Set oRATES = New MNBRate
oRX.Pattern = DATE_RX
oRX.Global = True	
oRATES.Init "EUR;USD;PLN;SEK;DKK;NOK;CHF;CZK;RON","100;100;1;100;100;100;100;1;1",";"
oMAIL.AddAdmin = "tomas.ac@volvo.com;tomas.chudik@volvo.com"
' Pre SAP dopice -> "EUR;USD;PLN;SEK;DKK;NOK;CHF;CZK;RON","100;100;1;100;100;100;100;1;1",";"

'Initialize TCD calendar
For Each tcd In Split(HU_TCD,";")
	oTCD.AddTCD tcd,""
Next

oTargetDate = CDate(Right("0000" & Year(Date),4) & "-" & Right("00" & Month(Date),2) & "-" & Right("00" & Day(Date),2))
'######################################
'############# M A I N ################
'######################################


'++++++++++++++++++++++++++++
'+ Create the output folder +
'++++++++++++++++++++++++++++
If Not oFSO.FolderExists(oFSO.GetParentFolderName(FOLDER)) Then
	oFSO.CreateFolder(oFSO.GetParentFolderName(FOLDER))
End If 
If Not oFSO.FolderExists(FOLDER) Then
	oFSO.CreateFolder(FOLDER)
End If
'+++++++++++++++++++++++++++
'+ Getting exchange rates  +
'+++++++++++++++++++++++++++
oTCD.FindNonTCDDate(oTargetDate) ' Find the first nonTCD day
debug.WriteLine "Target date: " & oDF.ToDayMonthYearWithDots(oTargetDate)
debug.WriteLine "Target date is a TCD: " & oTCD.IsTodayTCD
debug.WriteLine "First non TCD date: " & oDF.ToYearMonthDayWithDashes(oTCD.FirstNonTCD)


If Not oRATES.GetExRates(oDF.ToYearMonthDayWithDashes(oDF.ToYearMonthDayWithDashes(oTCD.FirstNonTCD))) Then

	Select Case oRATES.GetError
	
		Case 1
		
			debug.WriteLine oRATES.GetErrorMessage
			WScript.Quit(1)
			
		Case 2
			
			If Not oTCD.IsTodayTCD Then	
				debug.WriteLine oRATES.GetErrorMessage
				debug.WriteLine "Rates not published yet"
			Else
				debug.WriteLine oRATES.GetErrorMessage
			End If 
			WScript.Quit(2)
			
		Case Else 
		
	End Select 
	
End If 


If Not oTCD.IsTodayTCD And oTCD.FirstNonTCD <> oRATES.GetDate Then
	debug.WriteLine "First nonTCD day is TODAY, which means this date should be equal" _
				  & "to the date in the downloaded XML file." & vbCrLf & "First non TCD date: " _
				  & oDF.ToYearMonthDayWithDashes(oTCD.FirstNonTCD) & vbCrLf & "Date in the file: "
				  
	WScript.Quit(2)
End If 

debug.WriteLine "Target date in the XML file: " & oRATES.GetDate
debug.WriteLine "Calling MNBRate.MakeFile()"

If Not oRATES.MakeFile(FOLDER & "\" & oDF.ToYearMonthDay(oTargetDate) & ".txt") Then
	Debug.WriteLine "Error creating the output file"
	WScript.Quit(3)
End If 

'+++++++++++++++++++++++++++
'+  Upload exchange rates  +
'+++++++++++++++++++++++++++
Set oSAPLAUNCHER = New SAPLauncher
oSAPLAUNCHER.SetClientName = WScript.Arguments.Item(1)
oSAPLAUNCHER.SetSystemName = WScript.Arguments.Item(0)
oSAPLAUNCHER.SetLocalXML = oWSH.ExpandEnvironmentStrings("%APPDATA%") & "\SAP\Common\SAPUILandscape.xml"
oSAPLAUNCHER.CheckSAPLogon
oSAPLAUNCHER.FindSAPSession

If Not oSAPLAUNCHER.SessionFound Then
	oMAIL.SendMessage "Could not find the SAP session in the " & oSAPLAUNCHER.SAPsysName & ". Exiting.","E",oSAPLAUNCHER.SAPsysName
	WScript.Quit	
End If 

'All good. We have the session. continue
Set oSES = oSAPLAUNCHER.GetSession
oSES.StartTransaction "YTC_EXRATE_HU02"
oSES.findById("wnd[0]/usr/ctxtP_FILE").text = FOLDER & "\" & oDF.ToYearMonthDay(oTargetDate) & ".txt"
oSES.findById("wnd[0]/usr/txtP_KURST").text = "yhu1"
oSES.findById("wnd[0]/usr/ctxtP_GDATU").text = oDF.ToDayMonthYearWithDots(oTargetDate)
oSES.findById("wnd[0]").sendVKey 8
'condition handling of NOK status => every status bar text on wnd0 means error; should be send statusbar text to Error log 
oSES.findById("wnd[0]").sendVKey 0
'+++++++++++++++++++++++++++
'+  Verify uploaded rates  +
'+++++++++++++++++++++++++++
Set oRATECHECK = New RateCheck_v1
oRATECHECK.SAPSession = oSES
oRATECHECK.Init "HUF"
oRATECHECK.CheckRates FOLDER & "\" & oDF.ToYearMonthDay(oTargetDate) & ".txt", "yhu1"

If oRATECHECK.FilesVerified = 0 Then 

	debug.WriteLine file & " not found"
	oMAIL.SendMessage file & " not found","E",oSAPLAUNCHER.SAPsysName
	WScript.Quit
End If 

If oRATECHECK.GetNumIncompleteEntries > 0 Or oRATECHECK.GetNumInvalidEntries > 0 Or oRATECHECK.GetNumMissingEntries > 0 Then
	
	debug.WriteLine "Verification failed in " & oSAPLAUNCHER.SAPsysName
	oMAIL.SendMessage "Verification failed in " & oSAPLAUNCHER.SAPsysName & ". MISSING FILES: " & oRATECHECK.GetMissingEntries & " INCOMPLETE FILES: " & oRATECHECK.GetIncompleteEntries & " INVALID FILES: " & oRATECHECK.GetInvalidEntries,"E",oSAPLAUNCHER.SAPsysName
	WScript.Quit
	
Else	

	debug.WriteLine "Verification successfull in " & oSAPLAUNCHER.SAPsysName
	debug.WriteLine FOLDER & "\" & oDF.ToYearMonthDay(oTargetDate) & ".txt" & " is correct"
	oMAIL.SendMessage FOLDER & "\" & oDF.ToYearMonthDay(oTargetDate) & ".txt" & " successfully verified","I",oSAPLAUNCHER.SAPsysName
	
End If 

Checkin PROJECT, RESOURCE_NAME
WScript.Quit

'######################################
'########### M A I N   E N D  #########
'######################################





















'################################################
'###### C L A S S   D E F I N I T I O N S #######
'################################################
Class MNBRate

	Private strUrl
	Private oHTTP
	Private oXML
	Private oDATE
	Private nError
	Private sErrorMessage
	Private dictCurrs
	Private dictUnits
	
	Private Sub Class_Initialize()
		Set oXML = CreateObject("MSXML2.DOMDOCUMENT")
		Set oHTTP = CreateObject("MSXML2.XMLHTTP")
		Set dictUnits = CreateObject("Scripting.Dictionary")
		Set dictCurrs = CreateObject("Scripting.Dictionary")
		strUrl = "http://www.mnb.hu/arfolyamok.asmx"
	End Sub 
	
	Private Sub Class_Terminate()
	
	End Sub 
	
	Public Sub Init(strCurrencies,strUnits,chDelim)
		Dim aCurr
		Dim aUnit
		Dim i
		aCurr = Split(strCurrencies,chDelim)
		aUnit = Split(strUnits,chDelim)
		For i = 0 To UBound(aCurr)
			dictUnits.Add aCurr(i),CInt(aUnit(i))
		Next
	End Sub 
	
	Private Sub ResetError
		nError = 0
		sErrorMessage = ""
	End Sub 
	
	Private Sub SetError(nErrNum,sErrMsg)
		nError = nErrNum
		sErrorMessage = sErrMsg
	End Sub 
	
	Public Function MakeFile(strOutFilePath)
		ResetError
		Dim i
		Dim node
		Dim dRate
		Dim dEH ' EUR HUF
		Dim dSH ' SEK HUF
		Dim dES ' EUR SEK
		Dim oSTREAM
		Dim oFSO : Set oFSO = CreateObject("Scripting.FileSystemObject")
		Set oSTREAM = oFSO.OpenTextFile(strOutFilePath, ForWriting, True)
		If Not oFSO.FileExists(strOutFilePath) Then
			SetError 3,"File " & strOutFilePath & " not found"
			MakeFile = False
			Exit Function
		End If
		 
		For i = 0 To oXML.getElementsByTagName("Rate").length - 1
			Set node = oXML.getElementsByTagName("Rate").item(i)
			dictCurrs.Add node.attributes.getNamedItem("curr").text, CDbl(node.text)
		Next 
		
		For Each i In dictCurrs.Keys
			dRate = FormatNumber(Round(dictCurrs.Item(i) / dictUnits.Item(i),4),5) ' e.g EUR HUF
			
			If i = "EUR" Then 
				debug.WriteLine "Saving EUR HUF -> " & dRate
				dEH = dRate
			ElseIf i = "SEK" Then
				debug.WriteLine "Saving SEK HUF -> " & dRate
				dSH = dRate
			End If 
			
			If i = "PLN" Then
				oSTREAM.WriteLine i & vbTab & "HUF" & vbTab & dRate & vbTab & dictUnits.Item(i) / dictUnits.Item(i) & vbTab & dictUnits.Item(i)
				oSTREAM.WriteLine "HUF" & vbTab & i & vbTab & FormatNumber(Round(((1 / dRate) * 100),4),5) & vbTab & (dictUnits.Item(i) * 100) & vbTab & dictUnits.Item(i) / dictUnits.Item(i)
			Else 
				oSTREAM.WriteLine i & vbTab & "HUF" & vbTab & dRate & vbTab & dictUnits.Item(i) / dictUnits.Item(i) & vbTab & dictUnits.Item(i)
				oSTREAM.WriteLine "HUF" & vbTab & i & vbTab & FormatNumber(Round((1 / dRate),4),5) & vbTab & dictUnits.Item(i) & vbTab & dictUnits.Item(i) / dictUnits.Item(i)
			End If 
		Next
		
		dES = FormatNumber(Round(dEH / dSH,4),5) ' EUR SEK
		oSTREAM.WriteLine "EUR" & vbTab & "SEK" & vbTab & dES & vbTab & "1" & vbTab & "1"
		oSTREAM.WriteLine "SEK" & vbTab & "EUR" & vbTab & FormatNumber(Round(1 / dES,4),5) & vbTab & "1" & vbTab & "1"
		oSTREAM.Close
		
		MakeFile = True
	End Function 
	
	Public Function GetExRates(strDate_)
		ResetError
		Dim nChild,child
		Dim strRequestBody : strRequestBody = "<s:Envelope xmlns:s=""http://schemas.xmlsoap.org/soap/envelope/"">" _
												& "<s:Body><GetExchangeRates xmlns=""http://www.mnb.hu/webservices/""" _
												& " xmlns:i=""http://www.w3.org/2001/XMLSchema-instance""><startDate>" _
												& strDate_ & "</startDate><endDate>" & strDate_ & "</endDate><currencyNames>" _
												& "EUR,USD,PLN,SEK,DKK,NOK,CHF,CZK,RON</currencyNames></GetExchangeRates></s:Body></s:Envelope>"
		
	   With oHTTP 
			.open "POST", strUrl, False
			.setRequestHeader "Accept","application/xml"
			.setRequestHeader "SOAPAction", "http://www.mnb.hu/webservices/MNBArfolyamServiceSoap/GetExchangeRates"
			.setRequestHeader "Content-Type","text/xml"
			.setRequestHeader "Content-Length", Len(strRequestBody)
			.send strRequestBody
		End With	
		
		'DEBUG
		debug.WriteLine "Request body"
		debug.WriteLine "  " & strRequestBody
		debug.WriteLine "Response body"
		debug.WriteLine "  " & oHTTP.responseText
		'DEBUG
		
		If Not oHTTP.status = 200 Then
			nError = oHTTP.status
			sErrorMessage = oHTTP.responseText	
			GetExRates = False
			Exit Function
		Else
			oXML.loadXML oHTTP.responseText
			oXML.setProperty "SelectionNamespaces", "xmlns:s=""http://schemas.xmlsoap.org/soap/envelope/"""
			oXML.setProperty "SelectionNamespaces", "xmlns=""http://www.mnb.hu/webservices/"""
			oXML.setProperty "SelectionNamespaces", "xmlns:i=""http://www.w3.org/2001/XMLSchema-instance"""
			
			If oXML.getElementsByTagName("GetExchangeRatesResult").length = 0 Then
				nError = 1
				sErrorMessage = "No rates available"
				GetExRates = False
				Exit Function
			End If 
			
			debug.WriteLine(oXML.getElementsByTagName("GetExchangeRatesResult").item(0).text)
			oXML.loadXML oXML.getElementsByTagName("GetExchangeRatesResult").item(0).text ' load first child
			
			
			If oXML.getElementsByTagName("Rate").length = 0 Then
				nError = 2
				sErrorMessage = "No rates available" 
				GetExRates = False
				Exit Function
			End If 
			debug.WriteLine oXML.getElementsByTagName("Day").item(0).attributes.getNamedItem("date").text
			oDATE = CDate(oXML.getElementsByTagName("Day").item(0).attributes.getNamedItem("date").text)
			GetExRates = True
			Exit Function 
		End If 
	End Function
	
	Public Property Get GetDate
		GetDate = CDate(oXML.getElementsByTagName("Day").item(0).attributes.getNamedItem("date").text)
	End Property 
	
	Public Property Get GetError
		GetError = nError
	End Property
	
	Public Property Get GetErrorMessage
		GetErrorMessage = sErrorMessage
	End Property 
	

End Class 




Class TCDCalendar
	' ------------- Private members ----------- 
	Private t_dict ' Scripting.Dictionary that holds TCD entries
	Private t_len ' t_dict.Count
	Private t_isDateTCD ' this variable is set to True if current item in the dictionary is a TCD. Usefull during iterations through the dictionary
	Private t_FirstNonTCDDay
	Private t_dictExceptions ' this dictionary holds days that would normally be considered TCD 
	
	' ------------- Constructor ----------------------
	
	Private Sub Class_Initialize
		Set t_dict = CreateObject("Scripting.Dictionary")
		Set t_dictExceptions = CreateObject("Scripting.Dictionary")
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
	
	Public Function AddTCDException(ddmmyyyy)
		t_dictExceptions.Add ddmmyyyy,""
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



'Function for checkig-in to the watchdog list
Function Checkin(sProjectName,sResourceName)
	Dim sUserName,sUserSecret,sSiteUrl,sDomain,sTenantID,sClientID,sXDigest,sAccessToken,tmp,rxResult
	Dim oHTTP : Set oHTTP = CreateObject("MSXML2.ServerXMLHTTP.3.0")
	Dim oXML : Set oXML = CreateObject("MSXML2.DOMDOCUMENT")
	Dim oCON : Set oCON = CreateObject("Adodb.Connection")
	Dim oRST : Set oRST = CreateObject("Adodb.Recordset")
	Dim oRX : Set oRX = New RegExp
	Dim connectionString : connectionString = "Provider=Microsoft.ACE.OLEDB.12.0;WSS;IMEX=1;RetrieveIds=Yes;" & _
							 "DATABASE=https://volvogroup.sharepoint.com/sites/unit-rc-sk-bs-it/CREDENTIALS;" & _
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
        		"mailgot.it.volvo.net" 
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