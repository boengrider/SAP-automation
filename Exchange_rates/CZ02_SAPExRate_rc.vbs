Option Explicit 

Const PROJECT = "CZ02_SAPExRate"
Const CRFILE = "C:\!AUTO\CREDENTIALS\logins.txt"
Dim oDF,oTCD,oWSH,oFSO,oCNBRATE,oRATEUPLOAD,oSAPLAUNCHER,oRATECHECK,oLOG,oMAIL,system,client,ss,uploadedfiles

Set oLOG = New Logger
Set oWSH = CreateObject("Wscript.Shell")
Set oCNBRATE = New CNBRate 
Set oFSO = CreateObject("Scripting.FileSystemObject")
Set oDF = New DateFormatter
Set oTCD = New TCDCalendar
Set oMAIL = New Mailer 
system = WScript.Arguments.Item(0)
client = WScript.Arguments.Item(1)

If Hour(Time()) < 14 Then
	oLOG.LogEvent "Time less than 14:00","INF",False
	debug.WriteLine "Hour less than 14:00"
	WScript.Quit 
End If

oLOG.LocalLogFile = oWSH.ExpandEnvironmentStrings("%SYSTEMDRIVE%") & "\ExRate"
oMAIL.AddAdmin = "tomas.ac@volvo.com,tomas.chudik@volvo.com"

oTCD.Init ' Initialize the new instance
oTCD.AddTCD "01012021","Novy rok"
oTCD.AddTCD "02042021","Velky piatok"
oTCD.AddTCD "05042021","Velkonocny pondelok"
oTCD.AddTCD "01052021","Sviatok prace"
oTCD.AddTCD "08052021","Den oslobodenia od fasizmu"
oTCD.AddTCD "05072021","Cyril a Metod"
oTCD.AddTCD "06072021","Upalenie J. Husa"
oTCD.AddTCD "28092021","Den ceskej statnosti"
oTCD.AddTCD "28102021","Den vzniku sam. cs statu"
oTCD.AddTCD "17112021","Den boja za slobodu a demokraciu"
oTCD.AddTCD "24122021","Stedry den"
oTCD.AddTCD "25122021","1. sv. vianocny"
oTCD.AddTCD "26122021","2. sv. vianocny"
oTCD.AddTCD "01012022","Novy rok"
oTCD.AddTCD "15042022","Velky piatok"
oTCD.AddTCD "18042022","Velkonocny pondelok"
oTCD.AddTCD "01052022","Sviatok prace"
oTCD.AddTCD "08052022","Den oslobodenia od fasizmu"
oTCD.AddTCD "05072022","Cyril a Metod"
oTCD.AddTCD "06072022","Upalenie J. Husa"
oTCD.AddTCD "28092022","Den ceskej statnosti"
oTCD.AddTCD "28102022","Den vzniku sam. cs statu"
oTCD.AddTCD "17112022","Den boja za slobodu a demokraciu"
oTCD.AddTCD "24122022","Stedry den"
oTCD.AddTCD "25122022","1. sv. vianocny"
oTCD.AddTCD "26122022","2. sv. vianocny"
oTCD.AddTCD "01012023","Novy rok"
oTCD.FindNonTCDDate(Date())

debug.WriteLine "Today is a TCD: " & CStr(oTCD.IsTodayTCD)
debug.WriteLine "First non TCD date: " & oTCD.FirstNonTCD

'Rate object initialization
'oCNBRATE.Init oWSH.ExpandEnvironmentStrings("%TEMP%") & "\" & CreateGUID & ".txt",Null,Null,"17.08.2022","C:\ExRate\CZ02"
oCNBRATE.Init oWSH.ExpandEnvironmentStrings("%TEMP%") & "\" & CreateGUID & ".txt",Null,Null,oDF.ToDayMonthYearWithDots(oTCD.FirstNonTCD),"C:\ExRate\CZ02"
'oCNBRATE.Fcurrs = "EUR,SEK,USD,PLN,HUF,BGN,HRK,DKK,GBP,NOK,RON,RUB"
oCNBRATE.Fcurrs = "EUR,SEK,USD,PLN,HUF,BGN,DKK,GBP,NOK,RON"
oCNBRATE.OverrideQuantity = "EUR,USD,PLN,DKK,GBP,NOK" ' Applies for home currency e.g. EUR CZK 0,26635 1 100 ; CZK EUR 3,756 100 1
'oCNBRATE.AdjustRate = "EUR,USD,PLN,HUF,DKK,GBP,NOK,RUB" 
oCNBRATE.AdjustRate = "EUR,USD,PLN,HUF,DKK,GBP,NOK" 
oCNBRATE.AddOutputDirs = "\\vcn.ds.volvo.net\cli-sd\sd1294\046629\output\02_CZ02_ExRateProcessing"


oLOG.LogEvent "Today is a TCD: " & CStr(oTCD.IsTodayTCD),"INF",False
oLOG.LogEvent "Donwloading rate for: " & oTCD.FirstNonTCD,"INF",False

'Download rate from ECB (txt file)
Select Case (oCNBRATE.CreateTEMPFile)

	Case 0
		oLOG.LogEvent "TEMP file created successfully for the system " & system,"INF",False
		
	Case -1
		oLOG.LogEvent "Error creating a temp file for the system " & system & ": " & oCNBRATE.ErrorCode,"ERR",False
		oMAIL.SendMessage "Error creating a temp file for the system " & system & ": " & oCNBRATE.ErrorCode,"E",system
		oLOG.ReleaseLogs
		WScript.Quit
		
End Select 

'Parse the previously downloaded rate file and create file suitable for SAP
Select Case (oCNBRATE.ParseRateFile)

	Case 0
		oLOG.LogEvent "Rate file parsed successfully for the system" & system,"INF",False
		
	Case -1
	
		Select Case oCNBRATE.ErrorCode
		
			Case 2
				oLOG.LogEvent "Error parsing rate file. Error code: " & oCNBRATE.ErrorCode & ". Dates don't match","ERR",False
				oMAIL.SendMessage "Error parsing rate file. Error code: " & oCNBRATE.ErrorCode & ". Dates don't match","E",system
				oLOG.ReleaseLogs
				WScript.Quit
				
			Case 1
				oLOG.LogEvent "Error parsing rate file. Error code: " & oCNBRATE.ErrorCode & ". Can't determine the target date","ERR",False
				oMAIL.SendMessage "Error parsing rate file. Error code: " & oCNBRATE.ErrorCode & ". Can't determine the target date","E",system
				oLOG.ReleaseLogs
				WScript.Quit
				
		End Select
		
End Select


Select Case (oCNBRATE.MakeOutputFile)

	Case 0
		oLOG.LogEvent "Output file created successfully for the system " & system,"INF",False
		
	Case -1
		oLOG.LogEvent "Error creating the output file for the system " & system & ": " & oCNBRATE.ErrorCode,"ERR",False
		oMAIL.SendMessage "Error creating the output file for the system " & system & ": " & oCNBRATE.ErrorCode,"E",system
		oLOG.ReleaseLogs
		
End Select


If oCNBRATE.OutputFiles = "" Then
	oLOG.LogEvent "No files were processed. Won't run the upload script","ERR",False
	oMAIL.SendMessage "No files were processed. Won't run the upload script","E",system
	oLOG.ReleaseLogs
	WScript.Quit
End If 

'**************************************
'******* Upload rates to SAP **********
'**************************************
Set oSAPLAUNCHER = New SAPLauncher
oSAPLAUNCHER.SetClientName = client
oSAPLAUNCHER.SetSystemName = system
oSAPLAUNCHER.SetLocalXML = oWSH.ExpandEnvironmentStrings("%APPDATA%") & "\SAP\Common\SAPUILandscape.xml"
oSAPLAUNCHER.CheckSAPLogon
oSAPLAUNCHER.FindSAPSession

If Not oSAPLAUNCHER.SessionFound Then
	oLOG.LogEvent "Could not find the SAP session in the " & oSAPLAUNCHER.SAPsysName & ". Exiting.","ERR",False
	oMAIL.SendMessage "Could not find the SAP session in the " & oSAPLAUNCHER.SAPsysName & ". Exiting.","E",oSAPLAUNCHER.SAPsysName
	WScript.Quit	
End If 

'All good. We have the session. continue
Set ss = oSAPLAUNCHER.GetSession
Set oRATEUPLOAD = New RateUpload_v1
oRATEUPLOAD.SAPSession = ss
'Upload the rate file
uploadedfiles = oRATEUPLOAD.UploadRates(oCNBRATE.OutputFiles,"ycz1")
'************************************
'******* Check uploaded rates *********
'**************************************
Set oRATECHECK = New RateCheck_v1
oRATECHECK.Init "CZK"
oRATECHECK.SAPSession = ss
oRATECHECK.CheckRates oCNBRATE.OutputFiles,"ycz1"

If oRATECHECK.FilesVerified = 0 Then 
	oMAIL.SendMessage oCNBRATE.OutputFiles & " not found","E",oSAPLAUNCHER.SAPsysName
	WScript.Quit
End If 

If oRATECHECK.GetNumIncompleteEntries > 0 Or oRATECHECK.GetNumInvalidEntries > 0 Or oRATECHECK.GetNumMissingEntries > 0 Then
	oMAIL.SendMessage "Verification failed in " & oSAPLAUNCHER.SAPsysName & ". MISSING FILES: " & oRATECHECK.GetMissingEntries & " INCOMPLETE FILES: " & oRATECHECK.GetIncompleteEntries & " INVALID FILES: " & oRATECHECK.GetInvalidEntries,"E",oSAPLAUNCHER.SAPsysName
	WScript.Quit
End If
 	

oMAIL.SendMessage oCNBRATE.OutputFiles & " successfully verified","I",oSAPLAUNCHER.SAPsysName
oLOG.ReleaseLogs
Checkin PROJECT, CRFILE
'##########################################
'############ M A I N   E N D #############
'##########################################




Function ClearIECache
	Dim shell
	Set shell = CreateObject("Wscript.Shell")
	shell.Run "RunDLL32.exe InetCpl.cpl,ClearMyTracksByProcess 8",0,True
End Function 

	  



Function CreateGUID
  Dim TypeLib
  Set TypeLib = CreateObject("Scriptlet.TypeLib")
  CreateGUID = Mid(TypeLib.Guid, 2, 36)
End Function


Class CNBRate
	Private oWSH
	Private strPathToOutputDir
	Private strDate
	Private oHTTP
	Private dictRate
	Private dictOverrideQuantity
	Private oOutFile
	Private oTempFile
	Private oFSO
	Private strTCurr
	Private strAdditionalOutputDirs ' Copy output file(s) here
	Private strOutputFileName
	Private strPathToTempFile
	Private strRateURL
	Private strFCurrList
	Private boolMakeOutputFile
	Private errno
	Private dictAdjustRate
	Private listOutputFiles ' Pass this to the uploading script
	
	' ---------- Class Constructor ------------
	Sub Class_Initialize
		Set oWSH = CreateObject("Wscript.Shell")
		Set oHTTP = CreateObject("MSXML2.XMLHTTP")
		Set oFSO = CreateObject("Scripting.FileSystemObject")
		Set dictRate = CreateObject("Scripting.Dictionary")
		Set dictAdjustRate = CreateObject("Scripting.Dictionary")
		Set dictOverrideQuantity = CreateObject("Scripting.Dictionary")
		strRateURL = "https://www.cnb.cz/cs/financni-trhy/devizovy-trh/kurzy-devizoveho-trhu/kurzy-devizoveho-trhu/denni_kurz.txt"
		strTCurr = "CZK"
		boolMakeOutputFile = False 
		strAdditionalOutputDirs = Null 
		strPathToOutputDir = Null
		strFCurrList = Null
		strDate = Null
	End Sub 
	
	' ---------- Class Destructor -------------
	Sub Class_Terminate
		
	End Sub 
	
	' -------------------------------------------
	' --------- P u b l i c   M e t h o d s -----
	' -------------------------------------------
	
	' CreateTEMPFile
	' Method creates a unique temp file. 
	' If successfull it returns 0
	' On error it returns -1 and errno is set appropriately
	Public Function CreateTEMPFile
		If IsEmpty(strPathToTempFile) Or strPathToTempFile = "" Then
			boolMakeOutputFile = False 
			errno = 1 ' Can't create a temp file. No path to temp file was provided
			CreateTEMPFile = -1 ' No path provided. File cannot be created
			Exit Function
		End If
		
		Set oTempFile = oFSO.OpenTextFile(strPathToTempFile,2,True)
		ClearIECache ' Clear the IE cache first
		oHTTP.open "GET",strRateURL & "?date=" & strDate,False ' Open a http connection to the cnb server
		oHTTP.send  ' send request
		
		
		' In case URL cannot be accessed, close the temp file and delete it. Return HTTP error code
		If oHTTP.status <> 200 Then
			oTempFile.Close
			oFSO.DeleteFile strPathToTempFile
			boolMakeOutputFile = False
			errno = 2 ' Rate file can't be downloaded from the CNB web
			CreateTEMPFile = -1 ' ERROR downloading rate file
			Exit Function
		End If 
		' Write content to the temp file
		oTempFile.Write oHTTP.responseText
		oTempFile.Close
		boolMakeOutputFile = True
		CreateTEMPFile = 0 ' File created successfully
	End Function 
	
	Public Function DeleteTEMPFile
		If Not IsNull(oTempFile) And oFSO.FileExists(strPathToTempFile) Then
			oTempFile.Close
			oFSO.DeleteFile strPathToTempFile
		End If 
	End Function
	
	' Init()
	' strPath -> Absolute path to the temp file
	' strUrl  -> Rate URL or Null
	' strFCurrs -> comma delimited list of the wanted currencies or Null
	' strTargetDate -> target date or Null
	' strPathToOD -> Absolute path to the directory where to put otput files locally e.g C:\ExRate\CZ02
	Public Function Init(strPath,strUrl,strFCurrs,strTargetDate,strPathToOD) 		
		strPathToTempFile = strPath
		If Not IsNull(strPathToOD) Then
			strPathToOutputDir = strPathToOD
		Else ' Try to recover and use C:\ExRate path
				strPathToOutputDir = oWSH.ExpandEnvironmentStrings("%SYSTEMDRIVE%") & "\ExRate"
		End If
	
		Call MakeOutputDir ' Create utput directory
		
		
		If Not IsNull(strTargetDate) Then
			strDate = strTargetDate
		End If 
		
		If strUrl = "" Or IsNull(strUrl) Then
			errno = 1 ' No URL provided to the Init function
			Init = -1
			Exit Function
		Else 
			strRateURL = strUrl
		End If 
		
		If Not strFCurrs = "" Or Not IsNull(strFCurrs) Then
			strFCurrList = strFCurrs
		Else
			' Nothing
		End If 
			
		
	End Function 
	
	' ----------- ParseRateFile() ------------------------
	Public Function ParseRateFile() ' Call this method to add an entry into the dictionary. This function/method will do the parsing. It needs txt file with rates
		Dim key,column,line
	
		
		If IsNull(strDate) Then
			boolMakeOutputFile = False
			errno = 1 ' No Date was provided to the ParseRateFile() function
			ParseRateFile = -1 
			Exit Function
		End If 
   		
   		Set oTempFile = oFSO.OpenTextFile(strPathToTempFile,1,False)
   		line = oTempFile.ReadLine ' 1st line should contain date in DD.MM.YYYY  Compare it to our strDate
   		If Not strDate = Mid(line,1,10) Then
   			oTempFile.Close
   			If oFSO.FileExists(strPathToTempFile) Then
   				oFSO.DeleteFile strPathToTempFile
   			End If 
   			boolMakeOutputFile = False
   			errno = 2 ' Bad date. Target date and downloaded file date dont match
   			ParseRateFile = -1
   			Exit Function 
   		End If 
   		' ------ All OK continue parsing
   		Do While Not oTempFile.AtEndOfStream
   			
   			line = oTempFile.ReadLine ' 1st line should contain date in DD.MM.YYYY  Compare it to our strDate
   			column = Split(line,"|")
   			
   			For Each key In Split(strFCurrList,",")
   				
   				If key = column(3) Then 
  					
   					dictRate.Add column(3),column(4) ' RATE VALUE e.g. AUD 16,421
   						
   				End If 
   					
   			Next
   			
   		Loop
   		boolMakeOutputFile = True
   	End Function
   	
   	
   	Function MakeOutputFile
   		Dim key
   		' Check if the OD exists
   		If Not boolMakeOutputFile Then
   			errno = 1 ' Cant continue making output file
   			MakeOutputFile = -1
   			Exit Function
   		End If 
   		
   		If Not IsNull(strPathToOutputDir) Then
   			
   			strOutputFileName = Right("0000" & Year(Date() + 1),4) & Right("00" & Month(Date() + 1),2) & Right("00" & Day(Date() + 1),2) & ".txt"
   			Set oOutFile = oFSO.OpenTextFile(strPathToOutputDir & "\" & strOutputFileName,2,True)
   			For Each key In Split(strFCurrList,",")
 	
   				' First write all rates FOREIGN to HOME currency 
   				If dictAdjustRate.Exists(key) Then  ' If key exists we divide rate by 100
   					
   					If dictOverrideQuantity.Exists(key) Then ' If key exists we override ratio in the output file too

   						
   						oOutFile.WriteLine key & vbTab & strTCurr & vbTab & FormatNumber((dictRate.Item(key) / dictAdjustRate.Item(key)),5) & vbTab & "1" & vbTab & "100"
   							
   					Else
   						
   						oOutFile.WriteLine key & vbTab & strTCurr & vbTab & FormatNumber((dictRate.Item(key) / dictAdjustRate.Item(key)),5) & vbTab & "1" & vbTab & "1"
   							
   					End If
   					
   				Else   
   					
   					If dictOverrideQuantity.Exists(key) Then
   						
   						oOutFile.WriteLine key & vbTab & strTCurr & vbTab & FormatNumber(dictRate.Item(key),5) & vbTab & "1" & vbTab & "100"
   							
   					Else 
   							
   						oOutFile.WriteLine key & vbTab & strTCurr & vbTab & FormatNumber(dictRate.Item(key),5) & vbTab & "1" & vbTab & "1"
   							
   					End If 
   						
   				End If  
   				
   			Next 
   			
   			For Each key In Split(strFCurrList,",")	
   							
   				' Now write HOME to FOREIGN currency
   				' FormatNumber(Round((1/ORIGINAL RATE FROM FILE),5) * 100,5)
   				If dictAdjustRate.Exists(key) Then 
   				
   					If dictOverrideQuantity.Exists(key) Then
   																		
   						oOutFile.WriteLine strTCurr & vbTab & key & vbTab & FormatNumber((Round(1 / dictRate.Item(key),5) * dictAdjustRate.Item(key)),5) & vbTab & "100" & vbTab & "1"
   							
   					Else 
   						
   						oOutFile.WriteLine strTCurr & vbTab & key & vbTab & FormatNumber((Round(1 / dictRate.Item(key),5) * dictAdjustRate.Item(key)),5) & vbTab & "1" & vbTab & "1"
   							
   					End If 
   						
   						
   				Else  
   					
   					If dictOverrideQuantity.Exists(key) Then
   	
   						oOutFile.WriteLine strTCurr & vbTab & key & vbTab & FormatNumber(Round((1/dictRate.Item(key)),5),5) & vbTab & "1" & vbTab & "100"
   							
   					Else
   						
   						oOutFile.WriteLine strTCurr & vbTab & key & vbTab & FormatNumber(Round((1/dictRate.Item(key)),5),5) & vbTab & "1" & vbTab & "1"
   							
   					End If
   						
   						
   				End If  
   				
   				
   			Next
   			
   			oOutFile.WriteLine "EUR" & vbTab & "SEK" & vbTab & FormatNumber(Round((dictRate("EUR") / dictRate("SEK")),5),5) & vbTab & "1" & vbTab & "1"
   			oOutFile.WriteLine "SEK" & vbTab & "EUR" & vbTab & FormatNumber(Round((1 / ((dictRate("EUR") / dictRate("SEK")))),5),5) & vbTab &  "1" & vbTab & "1"
   			oOutFile.Close
   			oTempFile.Close
   			DeleteTEMPFile
   			listOutputFiles = strPathToOutputDir & "\" & strOutputFileName & ","
   			MakeOutputFile = 0
   			
   		End If 
   			
   		
   		
   	End Function 
   	
   	Private Function MakeOutputDir
	Dim comps,i,l,path 
	comps = Split(strPathToOutputDir,"\")
	l = UBound(comps) ' save len
	i = 0
	
	Do While Not i = l + 1
		
		If oFSO.GetDriveName(comps(i)) = comps(i) Then
			path = comps(i)
			i = i + 1
		End If
		
		path = path & "\" & comps(i)
		If Not oFSO.FolderExists(path) Then
			oFSO.CreateFolder path
		End If
		
		i = i + 1
		
	Loop 
	
End Function

   
   			
   	
   	' ----------------------------------------------
	' -------------- P r o p e r t i e s -----------
	' ----------------------------------------------
	Public Property Get RateURL
		RateURL = strRateURL
	End Property 
	
	Public Property Get PathToTempFile
		PathToTempFile = strPathToTempFile
	End Property
	
	Public Property Get Tcurr 
		Tcurr = strTCurr
	End Property 
	
	Public Property Let Tcurr(strCurr)
		strTCurr = strCurr
	End Property 
	
	Public Property Get Fcurrs
		Fcurrs = strFCurrList
	End Property 
	
	Public Property Let Fcurrs(strCurrs)
		strFCurrList = strCurrs
	End Property 
	
	Public Property Get Ddate 
		Ddate = strDate
	End Property 
	
	Public Property Get GetRate(strCurrency)
		GetRate = dictRate.Item(strCurrency)
	End Property 
	
	Public Property Let OverrideQuantity(strQuantity)
		Dim key
		For Each key In Split(strQuantity,",")
			dictOverrideQuantity.Add key,100
		Next
	End Property 
	
	Public Property Get ErrorCode
		ErrorCode = errno
	End Property 
	
	Public Property Let AdjustRate(strRates) ' comma delimited string
		Dim key
		For Each key In Split(strRates,",")
			dictAdjustRate.Add key,100
		Next
	End Property 
	
	Public Property Let AddOutputDirs(strDirs)
		strAdditionalOutputDirs = strDirs
	End Property 
	
	Public Property Get OutputFiles
		If listOutputFiles = "" Then 
			debug.WriteLine "listOutputFiles is empty"
			OutputFiles = ""
			Exit Property
		End If 
		OutputFiles = Left(listOutputFiles,Len(listOutputFiles) - 1)
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
	Public Function Init
		Set t_dict = CreateObject("Scripting.Dictionary")
		t_len = t_dict.Count
		t_FirstNonTCDDay = Null
		t_isDateTCD = False ' Initially false. GetLastNonTCDDate(D) uses this variable
	End Function
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



' ------------ Logger----------------

Class Logger
	
	Private boolRemoteLogExists		' Set to True if log file exists
	Private boolLocalLogExists		' Set to True if log file exists
	Private strUser					' String representing user
	Private strComputer				' String represnting computer
	Private strDate					' Current date
	Private strSource				' Script name 
	Private strLocalLog				' Path to the local log file
	Private strRemoteLog			' UNC path to the remote log file
	Private oFSO					' File system object
	Private oNET					' Network object. Used to obtain the user name
	Private dictSeverity			' Dictionary holding severity codes
	Private oLocalLog				' Local log file descriptor
	Private oRemoteLog				' Remote log file descriptor
	
	Sub Class_Initialize
	
		boolRemoteLogExists = False	' Initially set to False
		boolLocalLogExists = False	' Initially set to False
		Set oNET = CreateObject("Wscript.Network")
		Set oFSO = CreateObject("Scripting.FileSystemObject")
		Set dictSeverity = CreateObject("Scripting.Dictionary")
		dictSeverity.Add "ERR","E"
		dictSeverity.Add "INF","I"
		dictSeverity.Add "WRN","W"
		strUser = oNET.UserName
		strComputer = oNET.ComputerName
		strDate = Date()
		strSource = WScript.ScriptName
		strLocalLog = Null
		strRemoteLog = Null
		oLocalLog = Null
		oRemoteLog = Null
		
	End Sub 
	
	Sub Class_Terminate
		
		If Not IsNull(strLocalLog) Then
		
			oLocalLog.Close
		
		End If 
		
		If Not IsNull(strRemoteLog) Then
		
			oRemoteLog.Close
		
		End If 
		
	End Sub 
	
	'LogEvent method
	' strMessage -> Message to log
	' strSeverity -> Severity code, e.g "ERROR"
	' boolLogRemote -> If false log locally only. If true log remotely aswell
	Public Function LogEvent(strMessage,strSeverity,boolLogRemote)
		
		Select Case boolLogRemote
		
			Case True
			
				If Not IsNull(oLocalLog) And boolLocalLogExists Then
					oLocalLog.WriteLine strSource & vbTab & strDate & vbTab & Time() & vbTab & strUser & "@" & strComputer & vbTab & strMessage & vbTab & dictSeverity(strSeverity)
				End If 
				
				If Not IsNull(oRemoteLog) And boolRemoteLogExists Then
					oRemoteLog.WriteLine strSource & vbTab & strDate & vbTab & Time() & vbTab & strUser & "@" & strComputer & vbTab & strMessage & vbTab & dictSeverity(strSeverity)
				End If 
				
			Case False 
			
				If Not IsNull(oLocalLog) And boolLocalLogExists Then 
					oLocalLog.WriteLine strSource & vbTab & strDate & vbTab & Time() & vbTab & strUser & "@" & strComputer & vbTab & strMessage & vbTab & dictSeverity(strSeverity)
				End If 
				
		End Select 
	
	End Function 
	
	Public Function Header
	
		oLocalLog.WriteLine vbCrLf & "_____________________ " & WScript.ScriptName & " _____________________" & vbCrLf
		oRemoteLog.WriteLine vbCrLf &"_____________________ " & WScript.ScriptName & " _____________________" & vbCrLf 
		
	End Function 
	
	Public Function ReleaseLogs
		
		If Not IsNull(strLocalLog) Or Not strLocalLog = "" Then
		
			If IsObject(oLocalLog) Then 
				debug.WriteLine "Local log released"
				oLocalLog.Close
			End If 
		
		End If 
		
		If Not IsNull(strRemoteLog) Or Not strRemoteLog = "" Then
		
			If IsObject(oRemoteLog) Then 
				debug.WriteLine "Remote log released"
				oRemoteLog.Close
			End If 
		
		End If 
		
	End Function 
	
	
	

	' LocalLogFile property. 
	' Opens a log file for appending. If it doesn't exist, creates it
	Public Property Let LocalLogFile(strPath)
		
		If IsNull(strPath) Then
			boolLocalLogExists = False
			Exit Property
		End If 
		
		If oFSO.FolderExists(strPath) Then 
			Set oLocalLog = oFSO.OpenTextFile(strPath & "\log.txt",8,True)
			If oFSO.FileExists(strPath & "\log.txt") Then
				boolLocalLogExists = True
			End If 
		End If 
		
	End Property 
	
	' RemoteLogFile property. 
	' Opens a log file for appending. If it doesn't exist, creates it
	Public Property Let RemoteLogFile(strPath)
		
		If IsNull(strPath) Then
			boolRemoteLogExists = False
			Exit Property
		End If
		
		If oFSO.FolderExists(strPath) Then
			Set oRemoteLog = oFSO.OpenTextFile(strPath & "\log.txt",8,True)
			If oFSO.FileExists(strPath & "\log.txt") Then
				boolRemoteLogExists = True
			End If 
		End If 
		
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
		
		'Don't exit. This is now handled by rate check part
		'oSES.findById("wnd[0]/tbar[0]/okcd").text = "/NEX" ' Close transaction window
		'oSES.findById("wnd[0]").sendVKey 0
	
		UploadRates = i ' Return the number of uploaded files or 0 if error occured   

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
	Public Property Let SAPSession(ses)
		Set oSES = ses
	End Property 	
		

	
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

'Function for checkig-in to the watchdog list
Function Checkin(sProjectName,sCredFilePath)
	Dim sUserName,sUserSecret,sSiteUrl,sDomain,sTenantID,sClientID,sXDigest,sAccessToken,tmp,rxResult
	Dim oHTTP : Set oHTTP = CreateObject("MSXML2.ServerXMLHTTP.3.0")
	Dim oXML : Set oXML = CreateObject("MSXML2.DOMDocument")
	Dim oRX : Set oRX = New RegExp
	
	'Load credentials
	oXML.load sCredFilePath
	sUserName = oXML.selectSingleNode("//service[@name=""unit-rc-sk-bs-it""]/username").text
	sUserSecret = oXML.selectSingleNode("//service[@name=""unit-rc-sk-bs-it""]/password").text
	sSiteUrl = oXML.selectSingleNode("//service[@name=""unit-rc-sk-bs-it""]/host").text
	sDomain = oXML.selectSingleNode("//service[@name=""unit-rc-sk-bs-it""]/domain").text
	
	'Get TenantID & ClientID/ResourceID
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
