Option Explicit
'To test set TEST to true
Const PROJECT = "HR02_VIPSINT"
Const CRFILE = "C:\!AUTO\CREDENTIALS\logins.txt"
Const TEST = False
Const TESTDATE = "08. 05. 2022"
Const COUNTRY = "Croatia"
Const ROOT_DIRECTORY = "C:\!AUTO\HR02_VIPSINT"
Const CONFIGFILEPATH = "C:\!AUTO\CONFIGURATION\VIPSINT.conf"
Const CREDENTIALS = "C:\!AUTO\CREDENTIALS\logins.txt"
Const adSaveCreateNotExist = 1
Const adSaveCreateOverWrite = 2
Const adLockOptimistic = 3
Const adWriteChar = 0
Const adStateOpen = 1
Const adWriteLine = 1
Const ForAppending = 8
Const ForReading = 1
Const ForWriting = 2
Const adModeRead = 1
Const adModeReadWrite = 3
Const adModeRecursive = 4194304
Const adModeShareDenyNone = 16
Const adModeShareDenyRead = 4
Const adModeShareDenyWrite = 8
Const adModeShareExclusive = 12
Const adModeUnknown = 0
Const adModeWrite = 2
Dim PROCESS_SUPPLIER_CREDIT : PROCESS_SUPPLIER_CREDIT = True
Dim PROCESS_SUPPLIER : PROCESS_SUPPLIER = True
Dim PROCESS_CUSTOMER : PROCESS_CUSTOMER = True
Dim UPLOAD_TO_SAP : UPLOAD_TO_SAP = True	' TCH: On 2022-12-12 I have set to False due to the AP/APC currency change HRK->EUR and I need to manually secure csv value change till EURo conversion golive 2023-01-01. Afterwards I will change the values back to True and set also the local currency to EUR in Configuration
Dim wsh : Set wsh = CreateObject("Wscript.Shell")
Dim sw : Set sw = New StopWatch 
Dim oXML : Set oXML = CreateObject("msxml2.domdocument")
Dim oFSO : Set oFSO = CreateObject("scripting.filesystemobject")
Dim oRX : Set oRX = New RegExp 
oRX.Pattern = "wscript"
oRX.IgnoreCase = True 
Dim arrFeedback : arrFeedback = Array("Invalid date format","Nope, that's not correct date format either","Nope, bad format again","Can you read ?!","Are you kidding me ?!","I've had enough you twat !")
Dim numAttempts : numAttempts = 0
Dim boolUploadToSAPretval : boolUploadToSAPretval = False 
Dim boolAPcreated : boolAPcreated = False   ' Set to True if the file was created
Dim boolARcreated : boolARcreated = False   ' Set to True if the file was created
Dim boolAPCcreated : boolAPCcreated = False ' Set to True if the file was created
Dim boolARuploaded : boolARuploaded = False ' Set to True if the file was uploaded
Dim boolAPuploaded : boolAPuploaded = False ' Set to True if the file was uploaded
Dim boolAPCuploaded : boolAPCuploaded = False ' Set to True if the file was uploaded
Dim boolARsm35 : boolARsm35 = False ' Set to True if SM35 check OK
Dim boolAPsm35 : boolAPsm35 = False ' Set to True if SM35 check OK
Dim boolAPCsm35 : boolAPCsm35 = False ' Set to True if SM35 check OK
Dim boolIsInteractive : boolIsInteractive = False ' True if run from wscript
Dim boolProcessSupplierCredit : boolProcessSupplierCredit = False   ' This is for interactive run
Dim boolProcessSupplier : boolProcessSupplier = False				' This is for interactive run
Dim boolProcessCustomer : boolProcessCustomer = False				' This is for interactive run
Dim boolUploadToSAP : boolUploadToSAP = False 						' This is for interactive run
'Dim oPartners : Set oPartners = CreateObject("scripting.dictionary")  ' Dictionary holding only External partners customer/supplier. Index is partner number. e.g oEP.Item(28612) returns EC (external customer) or ES (external supplier)
   												       		           ' All information contained within this dictionary is based on query 0
Dim oUPartners : Set oUPartners = CreateObject("scripting.dictionary") ' Dictionary holding uknown/missing partners from the VIPSINT config
Dim oInvoices : Set oInvoices = CreateObject("scripting.dictionary")   ' Dictioniary holding invoice/credit note numbers. Each ivnovice is per partner i.e inovice number is the key and the parner id value
Dim oMAIL : Set oMAIL = New Mailer   ' Technical reports go here
Dim oMAILB : Set oMAILB = New Mailer ' Bussiness reports go here
Dim oDF : Set oDF = New DateFormatter
Dim oSP : Set oSP = New SP
Dim parma,tradingpartner,profitcenter,taxcode,amount,payterm ' Variables holding data during processing. Variables are reused in each loop during processing
oMAIL.AddAdmin = "tomas.ac@volvo.com;tomas.chudik@volvo.com"
oMAILB.AddAdmin = "josipa.protega@volvo.com"
Dim sDate,oDate ' Dates are set either from user input if interacive wscript, from user input if interactive cscript
Dim strSAPsys   ' SAP system name
Dim strSAPcli   ' SAP client name
Dim key,message
Dim retval,item
Dim oSchemafile
Dim ss,sl
Dim strBusinessReport
Dim strTechnicalReport
Dim strAttachment : strAttachment = wsh.SpecialFolders.Item("MyDocuments") & "\SAP\SAP GUI\wnd.jpg"
debug.WriteLine "#############################################"
debug.WriteLine " Script run: " & Date() & " " & Time()
debug.WriteLine "#############################################" 
'***********************************************************************************
'Parse command line arguments and check whether we're running from wscript or cscript
'***********************************************************************************
If oRX.Test(LCase(WScript.FullName)) Then ' wscript
	oRX.Pattern = "system32"
	If oRX.Test(LCase(WScript.FullName)) Then ' 64 bit
		wsh.Exec "C:\Windows\SysWOW64\wscript.exe """ & WScript.ScriptFullName & "" ' Restart in 32 bit
		WScript.Quit
	End If 
	boolIsInteractive = True 
	oRx.Pattern = "^(?:(?:31(\/|-|\.)(?:0?[13578]|1[02]))\1|" _
			& "(?:(?:29|30)(\/|-|\.)(?:0?[13-9]|1[0-2])\2))" _
			& "(?:(?:1[6-9]|[2-9]\d)?\d{2})$|^(?:29(\/|-|\.)0" _
			& "?2\3(?:(?:(?:1[6-9]|[2-9]\d)?(?:0[48]|[2468][048]" _
			& "|[13579][26])|(?:(?:16|[2468][048]|[3579][26])00)" _
			& ")))$|^(?:0?[1-9]|1\d|2[0-8])(\/|-|\.)(?:(?:0?[1-" _
			& "9])|(?:1[0-2]))\4(?:(?:1[6-9]|[2-9]\d)?\d{2})$"
			
	' Prompt for a date value
	sDate = InputBox("Valid date formats:" & vbCrLf _
				   & "DD.MM.YYYY" & vbCrLf _
	               & "DD/MM/YYYY" & vbCrLf _
	               & "DD-MM-YYYY","Enter date",oDF.ToDayMonthYearWithDots(Date))

	Do While Not CStr(oRx.Test(sDate))
		If numAttempts > UBound(arrFeedback) Or sDate = "" Then WScript.Quit
		sDate = InputBox("Valid date formats:" & vbCrLf _
					   & "DD.MM.YYYY" & vbCrLf _
		               & "DD/MM/YYYY" & vbCrLf _
		               & "DD-MM-YYYY" & vbCrLf & vbCrLf & arrFeedback(numAttempts),"Enter date","")
		numAttempts = numAttempts + 1
	Loop
	
	sDate = Replace(sDate,".","-")
	oDate = CDate(sDate)
	sDate = CDate(sDate)
	sDate = oDF.ToYearMonthDay(sDate)
	 
	'Read input values
	boolProcessSupplierCredit = InputBox("Process supplier credit ?" & vbCrLf & vbCrLf & "Yes or No","Process supplier credit","Yes")
	If LCase(boolProcessSupplierCredit) = "yes" Then
		PROCESS_SUPPLIER_CREDIT = True
	ElseIf LCase(boolProcessSupplierCredit) = "no" Then
		PROCESS_SUPPLIER_CREDIT = False
	ElseIf boolProcessSupplierCredit = "" Then
		PROCESS_SUPPLIER_CREDIT = False
	Else
		PROCESS_SUPPLIER_CREDIT = False
	End If
	
	boolProcessSupplier = InputBox("Process supplier ?" & vbCrLf & vbCrLf & "Yes or No","Process supplier","Yes")
	If LCase(boolProcessSupplier) = "yes" Then
		PROCESS_SUPPLIER = True
	ElseIf LCase(boolProcessSupplier) = "no" Then
		PROCESS_SUPPLIER = False
	ElseIf boolProcessSupplier = "" Then
		PROCESS_SUPPLIER = False
	Else
		PROCESS_SUPPLIER = False
	End If
	
	boolProcessCustomer = InputBox("Process customer ?" & vbCrLf & vbCrLf & "Yes or No","Process supplier","Yes")
	If LCase(boolProcessCustomer) = "yes" Then
		PROCESS_CUSTOMER = True
	ElseIf LCase(boolProcessCustomer) = "no" Then
		PROCESS_CUSTOMER = False
	ElseIf boolProcessCustomer = "" Then
		PROCESS_CUSTOMER = False
	Else
		PROCESS_CUSTOMER = False
	End If
	
	boolUploadToSAP = InputBox("Upload to SAP ?" & vbCrLf & vbCrLf & "Yes or No","Upload to SAP","No")
	If LCase(boolUploadToSAP) = "yes" Then
		UPLOAD_TO_SAP = True
	ElseIf LCase(boolUploadToSAP) = "no" Then
		UPLOAD_TO_SAP = False
	ElseIf boolUploadToSAP = "" Then
		UPLOAD_TO_SAP = False
	Else
		UPLOAD_TO_SAP = False
	End If
	
	If UPLOAD_TO_SAP Then 
		'Get SAP system name
		strSAPsys = InputBox("Enter SAP system name","SAP system name","FP2")
		If strSAPsys = "" Then WScript.Quit ' Quit silently. Can't continue without SAP system name
		
		'Get SAP client number
		strSAPcli = InputBox("Enter SAP client number","SAP client number","103")
		If strSAPcli = "" Then WScript.Quit ' Quit silently. Can't continue without SAP client number
	End If 
	
	'Verify input 
	If MsgBox("Please verify your input information before you continue" & vbCrLf & vbCrLf _
		 & "Date: " & oDF.ToDayMonthYearWithDots(oDate) & vbCrLf _
		 & "SAP system: " & strSAPsys & vbCrLf & "SAP client: " & strSAPcli & vbCrLf _
		 & "Process customer: " & CStr(boolProcessCustomer) & vbCrLf _
		 & "Process supplier: " & CStr(boolProcessSupplier) & vbCrLf _
		 & "Process supplier credit: " & CStr(boolProcessSupplierCredit) & vbCrLf _
		 & "Upload to SAP: " & CStr(boolUploadToSAP),vbInformation + vbYesNo,"Confirm action") = vbNo Then 
		 	WScript.Quit
	End If  ' End if user chooses No
		
	'One more confirmation
	If MsgBox("Are you 100% positive you want to continue ?" & vbCrLf & "There is no going back after you click yes",vbExclamation + vbYesNo,"Confirm action") = vbNo Then
		WScript.Quit
	End If 
	
	'CONTINUE after all prompts
Else ' Running from something esle than wscript, i guess it's cscript 
	oRX.Pattern = "system32"
	If oRX.Test(WScript.Path) Then 
		WScript.Echo "You need to run this script in 32 bit version of cscript"
		WScript.Quit
	End If 
	
	If WScript.Arguments.Count <> 2 Then 
		WScript.Echo "Usage: " & WScript.ScriptName & " system client"
		WScript.Quit
	End If 

	If TEST Then 
		debug.WriteLine "TEST RUN"
		oDate = CDate(TESTDATE)
		sDate = oDF.ToYearMonthDay(oDate)
	Else
		debug.WriteLine "Normal run"
		oDate = Date
		sDate = oDF.ToYearMonthDay(oDate)
	End If 
	strSAPsys = WScript.Arguments.Item(0) ' SAP system name
	strSAPcli = WScript.Arguments.Item(1) ' SAP client name
	'At this point continue processing from cli session
End If
'Verify folders/file structure
If Not oFSO.FolderExists(oFSO.GetParentFolderName(ROOT_DIRECTORY)) Then oFSO.CreateFolder(oFSO.GetParentFolderName(sWorkingDirectory))      ' Create C:\!AUTO
If Not oFSO.FolderExists(ROOT_DIRECTORY) Then oFSO.CreateFolder(ROOT_DIRECTORY) 													        ' Create C:\!AUTO\HR02_VIPSINT
If Not oFSO.FolderExists(ROOT_DIRECTORY & "\SOURCE") Then oFSO.CreateFolder(ROOT_DIRECTORY & "\SOURCE")         							' Create C:\!AUTO\HR02_VIPSINT\SOURCE
If Not oFSO.FolderExists(ROOT_DIRECTORY & "\PROCESSED") Then oFSO.CreateFolder(ROOT_DIRECTORY & "\PROCESSED")   							' Create C:\!AUTO\HR02_VIPSINT\PROCESSED
If Not oFSO.FileExists(CONFIGFILEPATH) Then WScript.Quit(100)      					   													    ' Config file is missing
If Not oFSO.FolderExists(oFSO.GetParentFolderName(CREDENTIALS)) Then WScript.Quit(101) 														' Credentials folder is missing
If Not oFSO.FileExists(CREDENTIALS) Then WScript.Quit(101)																					' Credentials file is missing					       														
'Process credentials XML file
oXML.load CREDENTIALS
Dim oCredentialsSubtree : Set oCredentialsSubtree = oXML.selectSingleNode("//service[@name='unit-rc-sk-bs-it']")
Dim sUsername : sUsername = oCredentialsSubtree.selectSingleNode("username").text
Dim sSecret : sSecret = oCredentialsSubtree.selectSingleNode("password").text
Dim sHost : sHost = oCredentialsSubtree.selectSingleNode("host").text
Dim sDomain : sDomain = oCredentialsSubtree.selectSingleNode("domain").text
'Verify credentials
If IsEmpty(sUsername) Or IsNull(sUsername) Or sUsername = "" Then WScript.Quit(201) End If  ' Can't continue without credentials
If IsEmpty(sSecret) Or IsNull(sSecret) Or sSecret = "" Then WScript.Quit(202) End If 		' Can't continue without credentials
If IsEmpty(sHost) Or IsNull(sHost) Or sHost = "" Then WScript.Quit(203) End If 		        ' Can't continue without host name/ip address
'Process configuration XML file
oXML.load CONFIGFILEPATH
Dim oConfigurationSubtree : Set oConfigurationSubtree = oXML.selectSingleNode("//Country[@name='" & COUNTRY & "']")
'Dim bProcessSupplierCredit : bProcessSupplierCredit = oConfigurationSubtree.selectSingleNode("//SupFlowCredit").attributes.getNamedItem("bool").text ' Process supplier credit
Dim sWorkingDirectory : sWorkingDirectory = oConfigurationSubtree.selectSingleNode("//WorkingDirectory").text					   ' Local working directcory
Dim sSharepointSourceDirectory : sSharepointSourceDirectory = oConfigurationSubtree.selectSingleNode("//SharepointSourceDirectory").text ' Sharepoint source directory
Dim sCurrency : sCurrency = oConfigurationSubtree.selectSingleNode("//Currency").text											   ' Currency
Dim sCompanyCode : sCompanyCode = oConfigurationSubtree.selectSingleNode("//CompanyCode").text									   ' SAP company code
Dim sCustRefSuffix : sCustRefSuffix = oConfigurationSubtree.selectSingleNode("//CustomerRefferenceSuffix").text					   ' Customer reference suffix
Dim sSessionAR : sSessionAR = oConfigurationSubtree.selectSingleNode("//SessionAR").text										   ' AR session name
Dim sSessionAP : sSessionAP = oConfigurationSubtree.selectSingleNode("//SessionAP").text										   ' AP session name
Dim sSessionAPC : sSessionAPC = oConfigurationSubtree.selectSingleNode("//SessionAPC").text										   ' APC session name
Dim sHeaderTextAR : sHeaderTextAR = oConfigurationSubtree.selectSingleNode("//DocHdrTxtAR").text								   ' AR header text
Dim sHeaderTextAP : sHeaderTextAP = oConfigurationSubtree.selectSingleNode("//DocHdrTxtAP").text								   ' AP header text
Dim sHeaderTextAPC : sHeaderTextAPC = oConfigurationSubtree.selectSingleNode("//DocHdrTxtAPC").text								   ' APC header text
Dim sDocTypeAR : sDocTypeAR = oConfigurationSubtree.selectSingleNode("//DocTypeAR").text								   		   ' AR doctype
Dim sDocTypeAP : sDocTypeAP = oConfigurationSubtree.selectSingleNode("//DocTypeAP").text								   		   ' AP doctype
Dim sDocTypeAPC : sDocTypeAPC = oConfigurationSubtree.selectSingleNode("//DocTypeAPC").text								   		   ' APC doctype
Dim sOutFileName : sOutFileName = sDate
Dim sOutFilePrefixAR : sOutFilePrefixAR = oConfigurationSubtree.selectSingleNode("FileMask[@type='AR' and @component='prefix']").text 	 ' Output filename prefix AR
Dim sOutFileSuffixAR : sOutFileSuffixAR = oConfigurationSubtree.selectSingleNode("FileMask[@type='AR' and @component='suffix']").text 	 ' Output filename suffix AR
Dim sOutFilePrefixAP : sOutFilePrefixAP = oConfigurationSubtree.selectSingleNode("FileMask[@type='AP' and @component='prefix']").text 	 ' Output filename prefix AP
Dim sOutFileSuffixAP : sOutFileSuffixAP = oConfigurationSubtree.selectSingleNode("FileMask[@type='AP' and @component='suffix']").text 	 ' Output filename suffix AP
Dim sOutFilePrefixAPC : sOutFilePrefixAPC = oConfigurationSubtree.selectSingleNode("FileMask[@type='APC' and @component='prefix']").text ' Output filename prefix APC
Dim sOutFileSuffixAPC : sOutFileSuffixAPC = oConfigurationSubtree.selectSingleNode("FileMask[@type='APC' and @component='suffix']").text ' Output filename suffix APC
Dim sInFilePrefix : sInFilePrefix = oConfigurationSubtree.selectSingleNode("InputFilePrefix").text
Dim sInFileSuffix : sInFileSuffix = oConfigurationSubtree.selectSingleNode("InputFileSuffix").text
Dim arrFilesToDownload : arrFilesToDownload = Array("A","B","C","D","E","I","J")
sw.Activate ' Start stopwatch
If oFSO.FileExists(sWorkingDirectory & "\SOURCE\" & sOutFilePrefixAR & sOutFileName & sOutFileSuffixAR) Then
	debug.WriteLine "Deleting residual file " & sWorkingDirectory & "\SOURCE\" & sOutFilePrefixAR & sOutFileName & sOutFileSuffixAR
	oFSO.DeleteFile sWorkingDirectory & "\SOURCE\" & sOutFilePrefixAR & sOutFileName & sOutFileSuffixAR
End If
If oFSO.FileExists(sWorkingDirectory & "\SOURCE\" & sOutFilePrefixAP & sOutFileName & sOutFileSuffixAP) Then
	debug.WriteLine "Deleting residual file " & sWorkingDirectory & "\SOURCE\" & sOutFilePrefixAP & sOutFileName & sOutFileSuffixAP
	oFSO.DeleteFile sWorkingDirectory & "\SOURCE\" & sOutFilePrefixAP & sOutFileName & sOutFileSuffixAP
End If 
If oFSO.FileExists(sWorkingDirectory & "\SOURCE\" & sOutFilePrefixAPC & sOutFileName & sOutFileSuffixAPC) Then
	debug.WriteLine "Deleting residual file " & sWorkingDirectory & "\SOURCE\" & sOutFilePrefixAPC & sOutFileName & sOutFileSuffixAPC
	oFSO.DeleteFile sWorkingDirectory & "\SOURCE\" & sOutFilePrefixAPC & sOutFileName & sOutFileSuffixAPC
End If 
'****************************
' Rebuild the Schema.ini file
'****************************
debug.WriteLine "Rebuilding 'Schema.ini' file"
Set oSchemafile = oFSO.OpenTextFile(sWorkingDirectory & "\SOURCE\Schema.ini",ForWriting,True)
For item = 0 To UBound(arrFilesToDownload)
	arrFilesToDownload(item) = sInFilePrefix & oDF.ToYearMonthDayWithDashes(oDate) & sInFileSuffix & arrFilesToDownload(item) & ".CSV"
	If LCase(Mid(arrFilesToDownload(item),Len(arrFilesToDownload(item)) - 4,1)) = "a" Then 
		oSchemafile.WriteLine "[" & arrFilesToDownload(item) & "]"
		oSchemafile.WriteLine "ColNameHeader=False"
		oSchemafile.WriteLine "Col1=Partner Text"
		oSchemafile.WriteLine "Col2=Invoice Text"
		oSchemafile.WriteLine "Col3=F3 Text"
		oSchemafile.WriteLine "Col4=InvoiceType Text"
		oSchemafile.WriteLine "Col5=InvoiceDate Text"
		oSchemafile.WriteLine "Col6=F6 Text"
		oSchemafile.WriteLine "Col7=Payterm Text"
		oSchemafile.WriteLine "Col8=F8 Text"
		oSchemafile.WriteLine "Col9=F9 Text"
		oSchemafile.WriteLine "Col10=F10 Text"
		oSchemafile.WriteLine "Col11=F11 Text"
		oSchemafile.WriteLine "Col12=F12 Text"
		oSchemafile.WriteLine "Col13=F13 Text"
		oSchemafile.WriteLine "Col14=InvoiceTotal Text"
		oSchemafile.WriteLine "Col15=F15 Text"
		oSchemafile.WriteLine "Col16=F16 Text"
		oSchemafile.WriteLine "Col17=F17 Text"
		oSchemafile.WriteLine "Col18=F18 Text"
		oSchemafile.WriteLine "Format=Delimited(;)"
		oSchemafile.WriteLine "DecimalSymbol=,"
	ElseIf LCase(Mid(arrFilesToDownload(item),Len(arrFilesToDownload(item)) - 4,1)) = "b" Then
		oSchemafile.WriteLine "[" & arrFilesToDownload(item) & "]"
		oSchemafile.WriteLine "ColNameHeader=False"
		oSchemafile.WriteLine "Col1=Partner Text"
		oSchemafile.WriteLine "Col2=Invoice Text"
		oSchemafile.WriteLine "Col3=F3 Text"
		oSchemafile.WriteLine "Col4=F4 Text"
		oSchemafile.WriteLine "Col5=F5 Text"
		oSchemafile.WriteLine "Col6=Taxcode Text"
		oSchemafile.WriteLine "Col7=F7 Text"
		oSchemafile.WriteLine "Col8=F8 Text"
		oSchemafile.WriteLine "Col9=Taxamount Text"
		oSchemafile.WriteLine "Col10=F10 Text"
		oSchemafile.WriteLine "Col11=F11 Text"
		oSchemafile.WriteLine "Col12=F12 Text"
		oSchemafile.WriteLine "Col13=F13 Text"
		oSchemafile.WriteLine "Format=Delimited(;)"
		oSchemafile.WriteLine "DecimalSymbol=,"
	ElseIf LCase(Mid(arrFilesToDownload(item),Len(arrFilesToDownload(item)) - 4,1)) = "c" Then
		oSchemafile.WriteLine "[" & arrFilesToDownload(item) & "]"
		oSchemafile.WriteLine "ColNameHeader=False"
		oSchemafile.WriteLine "Col1=Partner Text"
		oSchemafile.WriteLine "Col2=Invoice Text"
		oSchemafile.WriteLine "Col3=F3 Text"
		oSchemafile.WriteLine "Col4=F4 Text"
		oSchemafile.WriteLine "Col5=F5 Text"
		oSchemafile.WriteLine "Col6=F6 Text"
		oSchemafile.WriteLine "Col7=F7 Text"
		oSchemafile.WriteLine "Col8=F8 Text"
		oSchemafile.WriteLine "Col9=F9 Text"
		oSchemafile.WriteLine "Col10=ProductGroup Text"
		oSchemafile.WriteLine "Col11=NetValue Text"
		oSchemafile.WriteLine "Col12=Quantity Text"
		oSchemafile.WriteLine "Col13=F13 Text"
		oSchemafile.WriteLine "Col14=F14 Text"
		oSchemafile.WriteLine "Col15=F15 Text"
		oSchemafile.WriteLine "Col16=Taxcode Text"
		oSchemafile.WriteLine "Col17=F17 Text"
		oSchemafile.WriteLine "Col18=F18 Text"
		oSchemafile.WriteLine "Col19=F19 Text"
		oSchemafile.WriteLine "Col20=F20 Text"
		oSchemafile.WriteLine "Col21=F21 Text"
		oSchemafile.WriteLine "Col22=F22 Text"
		oSchemafile.WriteLine "Col23=F23 Text"
		oSchemafile.WriteLine "Col24=F24 Text"
		oSchemafile.WriteLine "Col25=F25 Text"
		oSchemafile.WriteLine "Col26=F26 Text"
		oSchemafile.WriteLine "Col27=F27 Text"
		oSchemafile.WriteLine "Col28=F28 Text"
		oSchemafile.WriteLine "Format=Delimited(;)"
		oSchemafile.WriteLine "DecimalSymbol=,"
	ElseIf LCase(Mid(arrFilesToDownload(item),Len(arrFilesToDownload(item)) - 4,1)) = "d" Then
		oSchemafile.WriteLine "[" & arrFilesToDownload(item) & "]"
		oSchemafile.WriteLine "ColNameHeader=False"
		oSchemafile.WriteLine "Col1=Partner Text"
		oSchemafile.WriteLine "Col2=PartnerType Text"
		oSchemafile.WriteLine "Col3=F3 Text"
		oSchemafile.WriteLine "Col4=Invoice Text"
		oSchemafile.WriteLine "Col5=InvoiceDate Text"
		oSchemafile.WriteLine "Col6=F6 Text"
		oSchemafile.WriteLine "Col7=F7 Text"
		oSchemafile.WriteLine "Col8=F8 Text"
		oSchemafile.WriteLine "Col9=F9 Text"
		oSchemafile.WriteLine "Col10=InvoiceTotal Text"
		oSchemafile.WriteLine "Col11=F11 Text"
		oSchemafile.WriteLine "Col12=F12 Text"
		oSchemafile.WriteLine "Col13=F13 Text"
		oSchemafile.WriteLine "Col14=F14 Text"
		oSchemafile.WriteLine "Col15=F15 Text"
		oSchemafile.WriteLine "Col16=F16 Text"
		oSchemafile.WriteLine "Col17=F17 Text"
		oSchemafile.WriteLine "Col18=F18 Text"
		oSchemafile.WriteLine "Col19=F19 Text"
		oSchemafile.WriteLine "Col20=F20 Text"
		oSchemafile.WriteLine "Col21=F21 Text"
		oSchemafile.WriteLine "Format=Delimited(;)"
		oSchemafile.WriteLine "DecimalSymbol=,"
	ElseIf LCase(Mid(arrFilesToDownload(item),Len(arrFilesToDownload(item)) - 4,1)) = "e" Then
		oSchemafile.WriteLine "[" & arrFilesToDownload(item) & "]"
		oSchemafile.WriteLine "ColNameHeader=False"
		oSchemafile.WriteLine "Col1=Partner Text"
		oSchemafile.WriteLine "Col2=F2 Text"
		oSchemafile.WriteLine "Col3=F3 Text"
		oSchemafile.WriteLine "Col4=Invoice Text"
		oSchemafile.WriteLine "Col5=InvoiceDate Text"
		oSchemafile.WriteLine "Col6=F6 Text"
		oSchemafile.WriteLine "Col7=F7 Text"
		oSchemafile.WriteLine "Col8=F8 Text"
		oSchemafile.WriteLine "Col9=F9 Text"
		oSchemafile.WriteLine "Col10=InvoiceTotal Text"
		oSchemafile.WriteLine "Col11=ProductGroup Text"
		oSchemafile.WriteLine "Col12=F12 Text"
		oSchemafile.WriteLine "Col13=Quantity Text"
		oSchemafile.WriteLine "Col14=F14 Text"
		oSchemafile.WriteLine "Col15=F15 Text"
		oSchemafile.WriteLine "Col16=Taxcode Text"
		oSchemafile.WriteLine "Col17=F17 Text"
		oSchemafile.WriteLine "Col18=NetValue Text"
		oSchemafile.WriteLine "Col19=F19 Text"
		oSchemafile.WriteLine "Col20=DealerNumber text"
		oSchemafile.WriteLine "Col21=F21 Text"
		oSchemafile.WriteLine "Col22=F22 Text"
		oSchemafile.WriteLine "Col23=F23 Text"
		oSchemafile.WriteLine "Col24=F24 Text"
		oSchemafile.WriteLine "Col25=F25 Text"
		oSchemafile.WriteLine "Col26=F26 Text"
		oSchemafile.WriteLine "Col27=F27 Text"
		oSchemafile.WriteLine "Col28=F28 Text"
		oSchemafile.WriteLine "Col29=F29 Text"
		oSchemafile.WriteLine "Format=Delimited(;)"
		oSchemafile.WriteLine "DecimalSymbol=,"
	ElseIf LCase(Mid(arrFilesToDownload(item),Len(arrFilesToDownload(item)) - 4,1)) = "i" Then
		oSchemafile.WriteLine "[" & arrFilesToDownload(item) & "]"
		oSchemafile.WriteLine "ColNameHeader=False"
		oSchemafile.WriteLine "Col1=Supplier Text"
		oSchemafile.WriteLine "Col2=SupplierType Text"
		oSchemafile.WriteLine "Col3=Dealer Text"
		oSchemafile.WriteLine "Col4=Creditnote Text"
		oSchemafile.WriteLine "Col5=InvoiceDate Text"
		oSchemafile.WriteLine "Col6=F6 Text"
		oSchemafile.WriteLine "Col7=F7 Text"
		oSchemafile.WriteLine "Col8=InvoiceTotal Text"
		oSchemafile.WriteLine "Col9=F9 Text"
		oSchemafile.WriteLine "Col10=F10 Text"
		oSchemafile.WriteLine "Col11=F11 Text"
		oSchemafile.WriteLine "Col12=F12 Text"
		oSchemafile.WriteLine "Col13=F13 Text"
		oSchemafile.WriteLine "Col14=F14 Text"
		oSchemafile.WriteLine "Col15=F15 Text"
		oSchemafile.WriteLine "Format=Delimited(;)"
		oSchemafile.WriteLine "DecimalSymbol=,"
	ElseIf LCase(Mid(arrFilesToDownload(item),Len(arrFilesToDownload(item)) - 4,1)) = "j" Then
		oSchemafile.WriteLine "[" & arrFilesToDownload(item) & "]"
		oSchemafile.WriteLine "ColNameHeader=False"
		oSchemafile.WriteLine "Col1=Supplier Text"
		oSchemafile.WriteLine "Col2=SupplierType Text"
		oSchemafile.WriteLine "Col3=Dealer Text"
		oSchemafile.WriteLine "Col4=Creditnote Text"
		oSchemafile.WriteLine "Col5=Invoice Text"
		oSchemafile.WriteLine "Col6=F6 Text"
		oSchemafile.WriteLine "Col7=F7 Text"
		oSchemafile.WriteLine "Col8=ProductGroup Text"
		oSchemafile.WriteLine "Col9=Quantity Text"
		oSchemafile.WriteLine "Col10=Netval Text"
		oSchemafile.WriteLine "Col11=F11 Text"
		oSchemafile.WriteLine "Col12=F12 Text"
		oSchemafile.WriteLine "Format=Delimited(;)"
		oSchemafile.WriteLine "DecimalSymbol=,"
	End If 
Next 
oSchemafile.Close
'***********************
'Connect to Sharepoint
'and download files
'in the source directory
'***********************
Download
debug.WriteLine "Downloading input files (" & UBound(arrFilesToDownload) + 1 & ") from sharepoint"
'*********************
'Process customer data
'*********************
If PROCESS_CUSTOMER Then 
	debug.WriteLine "Processing customer"
	'On Error Resume Next
	boolARcreated = ProcessCustomer(sWorkingDirectory & "\SOURCE\", arrFilesToDownload)
	debug.WriteLine "AR file created -> " & CStr(boolARcreated)
	If err.number >= 1000 And err.number < 2000 Then ' errors
		debug.WriteLine err.number & " " & err.Source & " " & err.Description
		ErrorMessage "<HTML><HEAD>" & sSharepointSourceDirectory & "</HEAD><BODY><p>" _
		 & "<span style=""color:red"">Error source</span>: " & err.Source _
		 & "<br><span style=""color:red"">Error code</span>: " & err.number _
		 & "<br><span style=""color:red"">Error description</span>: " & err.Description _
		 & "</p></BODY></HTML>","E;" & WScript.ScriptName & ";" & oDF.ToYearMonthDayWithDashes(Date),Null
		WScript.Quit(err.number)
	ElseIf err.number >= 2000 And err.number < 3000 Then ' info/notifications
		debug.WriteLine err.number & " " & err.Source & " " & err.Description
		ErrorMessage "<HTML><HEAD>" & sSharepointSourceDirectory & "</HEAD><BODY><p>" _
		 & "<span style=""color:red"">Info source</span>: " & err.Source _
		 & "<br><span style=""color:red"">Info code</span>: " & err.number _
		 & "<br><span style=""color:red"">Info description</span>: " & err.Description _
		 & "</p></BODY></HTML>","I;" & WScript.ScriptName & ";" & oDF.ToYearMonthDayWithDashes(Date),Null
		WScript.Quit(err.number)
	ElseIf err.number = 0 Then
	Else
		debug.WriteLine err.number & " " & err.Source & " " & err.Description
		ErrorMessage "<HTML><HEAD>" & sSharepointSourceDirectory & "</HEAD><BODY><p>" _
		 & "<span style=""color:red"">Error source</span>: " & err.Source _
		 & "<br><span style=""color:red"">Error code</span>: " & err.number _
		 & "<br><span style=""color:red"">Error description</span>: " & err.Description _
		 & "</p></BODY></HTML>","E;" & WScript.ScriptName & ";" & oDF.ToYearMonthDayWithDashes(Date),Null
		WScript.Quit(err.number)
	End If
	On Error GoTo 0
End If 
'*********************
'Process supplier data
'*********************
If PROCESS_SUPPLIER Then
	debug.WriteLine "Processing supplier"
	'On Error Resume Next
	boolAPcreated = ProcessSupplier(sWorkingDirectory & "\SOURCE\", arrFilesToDownload)
	debug.WriteLine "AP file created -> " & CStr(boolAPcreated)
	If err.number >= 1000 And err.number < 2000 Then ' errors
		debug.WriteLine err.number & " " & err.Source & " " & err.Description
		ErrorMessage "<HTML><HEAD>" & sSharepointSourceDirectory & "</HEAD><BODY><p>" _
		 & "<span style=""color:red"">Error source</span>: " & err.Source _
		 & "<br><span style=""color:red"">Error code</span>: " & err.number _
		 & "<br><span style=""color:red"">Error description</span>: " & err.Description _
		 & "</p></BODY></HTML>","E;" & WScript.ScriptName & ";" & oDF.ToYearMonthDayWithDashes(Date),Null
		WScript.Quit(err.number)
	ElseIf err.number >= 2000 And err.number < 3000 Then ' info/notifications
		debug.WriteLine err.number & " " & err.Source & " " & err.Description
		ErrorMessage "<HTML><HEAD>" & sSharepointSourceDirectory & "</HEAD><BODY><p>" _
		 & "<span style=""color:red"">Info source</span>: " & err.Source _
		 & "<br><span style=""color:red"">Info code</span>: " & err.number _
		 & "<br><span style=""color:red"">Info description</span>: " & err.Description _
		 & "</p></BODY></HTML>","I;" & WScript.ScriptName & ";" & oDF.ToYearMonthDayWithDashes(Date),Null
		WScript.Quit(err.number)
	ElseIf err.number = 0 Then
	Else
		debug.WriteLine err.number & " " & err.Source & " " & err.Description
		ErrorMessage "<HTML><HEAD>" & sSharepointSourceDirectory & "</HEAD><BODY><p>" _
		 & "<span style=""color:red"">Error source</span>: " & err.Source _
		 & "<br><span style=""color:red"">Error code</span>: " & err.number _
		 & "<br><span style=""color:red"">Error description</span>: " & err.Description _
		 & "</p></BODY></HTML>","E;" & WScript.ScriptName & ";" & oDF.ToYearMonthDayWithDashes(Date),Null
		WScript.Quit(err.number)
	End If
	On Error GoTo 0
End If 
'****************************
'Process supplier credit data
'****************************
If PROCESS_SUPPLIER_CREDIT Then
	debug.WriteLine "Processing supplier credit"
	'On Error Resume Next
	boolAPCcreated = ProcessSupplierCredit(sWorkingDirectory & "\SOURCE\", arrFilesToDownload)
	debug.WriteLine "APC file created -> " & CStr(boolAPCcreated)
	If err.number >= 1000 And err.number < 2000 Then ' errors
		debug.WriteLine err.number & " " & err.Source & " " & err.Description
		ErrorMessage "<HTML><HEAD>" & sSharepointSourceDirectory & "</HEAD><BODY><p>" _
		 & "<span style=""color:red"">Error source</span>: " & err.Source _
		 & "<br><span style=""color:red"">Error code</span>: " & err.number _
		 & "<br><span style=""color:red"">Error description</span>: " & err.Description _
		 & "</p></BODY></HTML>","E;" & WScript.ScriptName & ";" & oDF.ToYearMonthDayWithDashes(Date),Null
		WScript.Quit(err.number)
	ElseIf err.number >= 2000 And err.number < 3000 Then ' info/notifications
		debug.WriteLine err.number & " " & err.Source & " " & err.Description
		ErrorMessage "<HTML><HEAD>" & sSharepointSourceDirectory & "</HEAD><BODY><p>" _
		 & "<span style=""color:red"">Info source</span>: " & err.Source _
		 & "<br><span style=""color:red"">Info code</span>: " & err.number _
		 & "<br><span style=""color:red"">Info description</span>: " & err.Description _
		 & "</p></BODY></HTML>","I;" & WScript.ScriptName & ";" & oDF.ToYearMonthDayWithDashes(Date),owsh.SpecialFolders("MyDocuments") & "\SAP\SAP GUI\wnd.jpg"
		WScript.Quit(err.number)
	ElseIf err.number = 0 Then
	Else
		debug.WriteLine err.number & " " & err.Source & " " & err.Description
		ErrorMessage "<HTML><HEAD>" & sSharepointSourceDirectory & "</HEAD><BODY><p>" _
		 & "<span style=""color:red"">Error source</span>: " & err.Source _
		 & "<br><span style=""color:red"">Error code</span>: " & err.number _
		 & "<br><span style=""color:red"">Error description</span>: " & err.Description _
		 & "</p></BODY></HTML>","E;" & WScript.ScriptName & ";" & oDF.ToYearMonthDayWithDashes(Date),Null
		WScript.Quit(err.number)
	End If 	
	On Error GoTo 0
End If 
'**************************
'Upload output files to SAP
'**************************
If UPLOAD_TO_SAP Then
	If Not boolAPcreated And Not boolAPCcreated And Not boolARcreated Then
		debug.WriteLine "No output files were produced"
		oMAIL.SendMessage "<HTML><HEAD>" & sSharepointSourceDirectory & "</HEAD><BODY><p>" _
		 & "<span style=""color:red"">Info source</span>: VIPSINT.TechnicalReport" _
		 & "<br><span style=""color:red"">Info description</span>: No output files were produced for " & sDate _
		 & "</p></BODY></HTML>","I;" & WScript.ScriptName & ";" & strSAPsys & ";" & oDF.ToYearMonthDayWithDashes(Date)
		 
		 WScript.Sleep 500
		 
		 oMAILB.SendMessage "<HTML><HEAD>" & sSharepointSourceDirectory & "</HEAD><BODY><p>" _
		 & "<span style=""color:red"">Info source</span>: VIPSINT.BusinessReport" _
		 & "<br><span style=""color:red"">Info description</span>: No output files were produced for " & sDate _
		 & "</p></BODY></HTML>","I;" & WScript.ScriptName & ";" & strSAPsys & ";" & oDF.ToYearMonthDayWithDashes(Date)
		WScript.Quit
	End If 
	
	Set sl = New SAPLauncher
	sl.SetClientName = strSAPcli
	sl.SetSystemName = strSAPsys
	sl.SetLocalXML = wsh.ExpandEnvironmentStrings("%APPDATA%") & "\SAP\Common\SAPUILandscape.xml"
	sl.CheckSAPLogon
	sl.FindSAPSession
	
	If Not sl.SessionFound Then 
		ErrorMessage "<HTML><HEAD>" & sSharepointSourceDirectory & "</HEAD><BODY><p>" _
		 & "<span style=""color:red"">Error source</span>: SAPLauncher.FindSAPSession" _
		 & "<br><span style=""color:red"">Error code</span>: 1" _
		 & "<br><span style=""color:red"">Error description</span>: Session not found" _
		 & "</p></BODY></HTML>","E;" & WScript.ScriptName & ";" & oDF.ToYearMonthDayWithDashes(Date),Null
		WScript.Quit(1)
	End If 
	
	Set ss = sl.GetSession
	
	
	debug.WriteLine "******************** Uploading to SAP **********************"
	On Error Resume Next
	'*****************************
	' AR file upload
	'*****************************
	If boolARcreated Then 
		debug.WriteLine "Uploading file to SAP: " & sWorkingDirectory & "\SOURCE", sOutFilePrefixAR & sOutFileName & sOutFileSuffixAR
		UploadToSAP ss,sWorkingDirectory & "\SOURCE", sOutFilePrefixAR & sOutFileName & sOutFileSuffixAR
		
		If CheckError Then
			HandleError ErrorHanlderB(False,wsh.SpecialFolders("MyDocuments") & "\SAP\SAP GUI\wnd.jpg",err.number,err.Source,err.Description,ss)
			'HandleError False,wsh.SpecialFolders("MyDocuments") & "\SAP\SAP GUI\wnd.jpg"
			WScript.Quit
		End If 
		
		boolARuploaded = True
		oFSO.CopyFile sWorkingDirectory & "\SOURCE\" & sOutFilePrefixAR & sOutFileName & sOutFileSuffixAR, sWorkingDirectory & "\PROCESSED\" & sOutFilePrefixAR & sOutFileName & sOutFileSuffixAR, True
		oFSO.DeleteFile sWorkingDirectory & "\SOURCE\" & sOutFilePrefixAR & sOutFileName & sOutFileSuffixAR
	End If  
	
	'*****************************
	' AP file upload
	'*****************************
	If boolAPcreated Then
		debug.WriteLine "Uploading file to SAP: " & sWorkingDirectory & "\SOURCE", sOutFilePrefixAP & sOutFileName & sOutFileSuffixAP
		UploadToSAP ss,sWorkingDirectory & "\SOURCE", sOutFilePrefixAP & sOutFileName & sOutFileSuffixAP
		
		If CheckError Then
			HandleError ErrorHanlderB(False,wsh.SpecialFolders("MyDocuments") & "\SAP\SAP GUI\wnd.jpg",err.number,err.Source,err.Description,ss)
			'HandleError False,wsh.SpecialFolders("MyDocuments") & "\SAP\SAP GUI\wnd.jpg"
			WScript.Quit
		End If 
		
		boolAPuploaded = True 
		oFSO.CopyFile sWorkingDirectory & "\SOURCE\" & sOutFilePrefixAP & sOutFileName & sOutFileSuffixAP, sWorkingDirectory & "\PROCESSED\" & sOutFilePrefixAP & sOutFileName & sOutFileSuffixAP, True
		oFSO.DeleteFile sWorkingDirectory & "\SOURCE\" & sOutFilePrefixAP & sOutFileName & sOutFileSuffixAP
	End If 
	
	'*****************************
	' APC file upload
	'*****************************
	If boolAPCcreated Then
		debug.WriteLine "Uploading file to SAP: " & sWorkingDirectory & "\SOURCE", sOutFilePrefixAPC & sOutFileName & sOutFileSuffixAPC
		UploadToSAP ss,sWorkingDirectory & "\SOURCE", sOutFilePrefixAPC & sOutFileName & sOutFileSuffixAPC
		
		If CheckError Then
			HandleError ErrorHanlderB(False,wsh.SpecialFolders("MyDocuments") & "\SAP\SAP GUI\wnd.jpg",err.number,err.Source,err.Description,ss)
			'HandleError False,wsh.SpecialFolders("MyDocuments") & "\SAP\SAP GUI\wnd.jpg"
			WScript.Quit
		End If
		
		boolAPCuploaded = True 
		oFSO.CopyFile sWorkingDirectory & "\SOURCE\" & sOutFilePrefixAPC & sOutFileName & sOutFileSuffixAPC, sWorkingDirectory & "\PROCESSED\" & sOutFilePrefixAPC & sOutFileName & sOutFileSuffixAPC, True
		oFSO.DeleteFile sWorkingDirectory & "\SOURCE\" & sOutFilePrefixAPC & sOutFileName & sOutFileSuffixAPC
	End If 								
	
	'*****************************
	' SM35 check
	'*****************************
	debug.WriteLine "******************** SM35 check **********************"
	If boolARuploaded Then
		debug.WriteLine "Checking " & sSessionAR 
		boolARsm35 = SM35(ss,sSessionAR,oDF.ToDayMonthYearWithDots(Date),wsh.ExpandEnvironmentStrings("%USERNAME%"),5,5000)
		debug.WriteLine " SM35 -> " & CStr(boolARsm35)
		
		
		If CheckError Then
			HandleError ErrorHanlderB(True,wsh.SpecialFolders("MyDocuments") & "\SAP\SAP GUI\wnd.jpg",err.number,err.Source,err.Description,ss)
		End If 
	End If 
	
	If boolAPuploaded Then 
		debug.WriteLine "Checking " & sSessionAP 
		boolAPsm35 = SM35(ss,sSessionAP,oDF.ToDayMonthYearWithDots(Date),wsh.ExpandEnvironmentStrings("%USERNAME%"),5,5000)
		debug.WriteLine " SM35 -> " & CStr(boolAPsm35)
		
		
		If CheckError Then
			HandleError ErrorHanlderB(True,wsh.SpecialFolders("MyDocuments") & "\SAP\SAP GUI\wnd.jpg",err.number,err.Source,err.Description,ss)
		End If 
	End If 
	
	If boolAPCuploaded Then 
		debug.WriteLine "Checking " & sSessionAPC 
		boolAPCsm35 = SM35(ss,sSessionAPC,oDF.ToDayMonthYearWithDots(Date),wsh.ExpandEnvironmentStrings("%USERNAME%"),5,5000)
		debug.WriteLine " SM35 -> " & CStr(boolAPCsm35)
	
		
		If CheckError Then
			HandleError ErrorHanlderB(True,wsh.SpecialFolders("MyDocuments") & "\SAP\SAP GUI\wnd.jpg",err.number,err.Source,err.Description,ss)
		End If 
	End If 
	
	ss.findById("wnd[0]/tbar[0]/okcd").text="/nex"
	ss.findById("wnd[0]").sendVKey 0
	On Error GoTo 0
End If 


sw.Deactivate
If boolIsInteractive And Not UPLOAD_TO_SAP Then 
	'MsgBox "Done",vbOK + vbInformation,"Operation completed"
	wsh.Run "explorer.exe """ & sWorkingDirectory & "\SOURCE",1,False
End If 




'*******************************
'Send the final technical and
'admin (business) reports
'*******************************
'Technical report
If Not boolIsInteractive Then 
	ErrorMessage "<!DOCTYPE html><HTML><HEAD>" & sSharepointSourceDirectory & "</HEAD>" _
			 & "<style>table, th, td { border:1px solid black; } th { background-color:#a1a1a1; }</style><BODY><p>" _
			 & "<span style=""color:red"">Info source</span>: VIPSINT.TechnicalReport</p>" _
			 & "<br>" & oDF.ToYearMonthDayWithDashes(oDate) & "<br>" _
			 & "<table><tr><th style=""text-align:left"">File name</th><th style=""text-align:left"">Create status</th><th style=""text-align:left"">Upload status</th><th style=""text-align:left"">SM35 status</th></tr>" _
			 & "<tr><td>" & sOutFilePrefixAR & sOutFileName & sOutFileSuffixAR & "</td><td>" & CStr(boolARcreated) & "</td><td>" & CStr(boolARuploaded) & "</td><td>" & CStr(boolARsm35) & "</td></tr>" _
			 & "<tr><td>" & sOutFilePrefixAP & sOutFileName & sOutFileSuffixAP & "</td><td>" & CStr(boolAPcreated) & "</td><td>" & CStr(boolAPuploaded) & "</td><td>" & CStr(boolAPsm35) & "</td></tr>" _
			 & "<tr><td>" & sOutFilePrefixAPC & sOutFileName & sOutFileSuffixAPC & "</td><td>" & CStr(boolAPCcreated) & "</td><td>" & CStr(boolAPCuploaded) & "</td><td>" & CStr(boolAPCsm35) & "</td></tr>" _
			 & "</table><br>" & strTechnicalReport & "</BODY></HTML>","I;" & WScript.ScriptName & ";" & strSAPsys & ";" & oDF.ToYearMonthDayWithDashes(Date) & ";" & sw.Duration,Null
End If 







'Business report only if upload to sap and script not run on demand
If UPLOAD_TO_SAP And Not boolIsInteractive Then
	If Not boolARcreated And Not boolAPcreated And Not boolAPCcreated Then 
		strBusinessReport = "<!DOCTYPE html><HTML><HEAD>" & sSharepointSourceDirectory & "</HEAD><BODY><p>" _
					  	 & "<span style=""color:red"">Info source</span>: VIPSINT.BusinessReport</p>" _
						 & "<br>" & oDF.ToYearMonthDayWithDashes(oDate) & "<br>" _
						 & "<p>No data for processing</p></BODY></HTML>"
	Else
		strBusinessReport = "<!DOCTYPE html><HTML><HEAD>" & sSharepointSourceDirectory & "</HEAD>" _
					     & "<style>table, th, td { border:1px solid black; } th { background-color:#a1a1a1; }</style><BODY><p>" _
					  	 & "<span style=""color:red"">Info source</span>: VIPSINT.BusinessReport</p>" _
						 & "<br>" & oDF.ToYearMonthDayWithDashes(oDate) & "<br>" _
						 & "<table><tr><th style=""text-align:left"">File name</th><th style=""text-align:left"">Upload status</th><th style=""text-align:left"">SM35 status</th></tr>"
		If boolARcreated And boolARuploaded Then
			strBusinessReport = strBusinessReport & "<tr><td>" & sOutFilePrefixAR & sOutFileName & sOutFileSuffixAR & "</td><td>Uploaded</td>"
			If boolARsm35 Then 
				strBusinessReport = strBusinessReport & "<td style=""color:#5CC22F"">Processed</td></tr>"
			Else 
				strBusinessReport = strBusinessReport & "<td style=""color:#FF8E00"">Check SAP</td></tr>"
			End If
		End If 
		
		If boolAPcreated And boolAPuploaded Then
			strBusinessReport = strBusinessReport & "<tr><td>" & sOutFilePrefixAP & sOutFileName & sOutFileSuffixAP & "</td><td>Uploaded</td>"
			If boolAPsm35 Then 
				strBusinessReport = strBusinessReport & "<td style=""color:#5CC22F"">Processed</td></tr>"
			Else 
				strBusinessReport = strBusinessReport & "<td style=""color:orange"">Check SAP</td></tr>"
			End If
		End If 
		
		If boolAPCcreated And boolAPCuploaded Then
			strBusinessReport = strBusinessReport & "<tr><td>" & sOutFilePrefixAPC & sOutFileName & sOutFileSuffixAPC & "</td><td>Uploaded</td>"
			If boolAPCsm35 Then 
				strBusinessReport = strBusinessReport & "<td style=""color:#5CC22F"">Processed</td></tr>"
			Else 
				strBusinessReport = strBusinessReport & "<td style=""color:orange"">Check SAP</td></tr>"
			End If
		End If 
		
		strBusinessReport = strBusinessReport & "</table></BODY></HTML>"
	End If 
		
	WScript.Sleep 500	
	oMAILB.SendMessage strBusinessReport,"I;" & WScript.ScriptName & ";" & strSAPsys & ";" & oDF.ToYearMonthDayWithDashes(Date)
End If 	 

'WdApp
Checkin PROJECT, CRFILE		 

'***************************************************************
'****************** M A I N   E N D ****************************
'***************************************************************

'**************************
'SM35 check function
'**************************
Function SM35(ByRef ses,strSessionName,strDate,strUser,numRoundTrips,numRoundTripDelay)
	Dim i
	sl.KillPopups(ses)
	If Not LCase(ses.Info.Transaction) = "sm35" Then	
		ses.findById("wnd[0]/tbar[0]/okcd").text = "/nSM35"
		ses.findById("wnd[0]").sendVKey 0
	End If 
	sl.KillPopups(ses)
	ses.findById("wnd[0]/usr/subD1000_HEADER:SAPMSBDC_CC:1005/txtD0100-MAPN").text = strSessionName
	ses.findById("wnd[0]/usr/subD1000_HEADER:SAPMSBDC_CC:1005/ctxtD0100-VON").text = strDate
	ses.findById("wnd[0]/usr/subD1000_HEADER:SAPMSBDC_CC:1005/ctxtD0100-BIS").text = strDate
	ses.findById("wnd[0]/usr/subD1000_HEADER:SAPMSBDC_CC:1005/txtD0100-CREATOR").text = strUser
	ses.findById("wnd[0]").sendVKey 0
	sl.KillPopups(ses)
	
	'At this point there should be at least one session found
	If CInt(ses.findById("wnd[0]/usr/subD1000_FOOT:SAPMSBDC_CC:1015/txtTC_APQI-LINES").text) = 0 Then
		strTechnicalReport = strTechnicalReport & "<p><span style=""color:red"">(SM35)</span> Expected at least one " & strSessionName & ". Found none</p>"
		Exit Function 
	End If 
	
	'Loop numRoundTrips for numRoundTripDelay
	For i = 1 To numRoundTrips
		ses.findById("wnd[0]").sendVKey 0
		If LCase(ses.findById("wnd[0]/usr/tabsD1000_TABSTRIP/tabpALLE/ssubD1000_SUBSCREEN:SAPMSBDC_CC:1010/tblSAPMSBDC_CCTC_APQI/lblITAB_APQI-STATUS[1,0]").tooltip) = "processed" Then
			SM35 = True
			Exit Function
		ElseIf LCase(ses.findById("wnd[0]/usr/tabsD1000_TABSTRIP/tabpALLE/ssubD1000_SUBSCREEN:SAPMSBDC_CC:1010/tblSAPMSBDC_CCTC_APQI/lblITAB_APQI-STATUS[1,0]").tooltip) = "errors" Then 
			SM35 = False
			Exit Function 
		End If
		WScript.Sleep numRoundTripDelay
	Next 
	
	'If after numRoundTrips document is still not processed append to technical report and exit function
	strTechnicalReport = strTechnicalReport & "<p><span style=""color:red"">(SM35)</span> Could not verify " & strSessionname & " after " & numRoundTrips & " retries</p>"
	SM35 = False 
	
End Function 

'*************************
'Function checking if global
'err obejct is set i.e. 
'exception occured
'**************************
Function CheckError()
	If err.number <> 0 Then 
		CheckError = True
	Else
		CheckError = False
	End If
End Function 

'**********************************
' General purpose exception handler
'**********************************
Function HandleError(ByRef fpointer)
	fpointer
End Function 

'*********************************
' ErrorHandlerB
'*********************************
Function ErrorHandlerB(boolTerminateB,strAttachmentPathB,numErrorCodeB,strErrorSourceB,strErrorDescriptionB,ByRef oSessionB)
	'Call ErrorHanlderA first
	ErrorHandlerA boolTerminateB,strAttachmentPathB,numErrorCodeB,strErrorSourceB,strErrorDescriptionB 
	'Additionally close SAP session so that it doesn't hang there
	oSessionB.findById("wnd[0]/tbar[0]/okcd").text="/nex"
	oSessionB.findById("wnd[0]").sendVKey 0
End Function 
'*********************************
' ErrorHandlerA 
'*********************************
Function ErrorHandlerA(boolTerminate,strAttachmentPath,numErrorCode,strErrorSource,strErrorDescription)
	If numErrorCode >= 1000 And numErrorCode < 2000 Then ' errors
		debug.WriteLine numErrorCode & " " & numErrorCode & " " & strErrorDescription
		ErrorMessage "<HTML><HEAD>" & sSharepointSourceDirectory & "</HEAD><BODY><p>" _
		 & "<span style=""color:red"">Error source</span>: " & strErrorSource _
		 & "<br><span style=""color:red"">Error code</span>: " & numErrorCode _
		 & "<br><span style=""color:red"">Error description</span>: " & strErrorDescription _
		 & "</p></BODY></HTML>","E;" & WScript.ScriptName & ";" & strSAPsys	& ";" & oDF.ToYearMonthDayWithDashes(Date),strAttachmentPath
		If boolTerminate Then WScript.Quit(numErrorCode)
	ElseIf numErrorCode >= 2000 And numErrorCode < 3000 Then ' info/notifications
		debug.WriteLine numErrorCode & " " & err.Source & " " & strErrorDescription
		ErrorMessage "<HTML><HEAD>" & sSharepointSourceDirectory & "</HEAD><BODY><p>" _
		 & "<span style=""color:red"">Info source</span>: " & strErrorSource _
		 & "<br><span style=""color:red"">Info code</span>: " & numErrorCode _
		 & "<br><span style=""color:red"">Info description</span>: " & strErrorDescription _
		 & "</p></BODY></HTML>","I;" & WScript.ScriptName & ";" & strSAPsys & ";" & oDF.ToYearMonthDayWithDashes(Date),strAttachmentPath
		If boolTerminate Then WScript.Quit(numErrorCode)
	ElseIf numErrorCode = 0 Then
	Else
		debug.WriteLine numErrorCode & " " & strErrorSource & " " & strErrorDescription
		ErrorMessage "<HTML><HEAD>" & sSharepointSourceDirectory & "</HEAD><BODY><p>" _
		 & "<span style=""color:red"">Error source</span>: " & strErrorSource _
		 & "<br><span style=""color:red"">Error code</span>: " & numErrorCode _
		 & "<br><span style=""color:red"">Error description</span>: " & strErrorDescription _
		 & "</p></BODY></HTML>","E;" & WScript.ScriptName & ";" & strSAPsys & ";" & oDF.ToYearMonthDayWithDashes(Date),strAttachmentPath
		If boolTerminate Then WScript.Quit(numErrorCode)
	End If
End Function 
	
'Function that handles 
'raised exception
'if boolTerminate is false
'script continues, otherwise
'script is terminated
'***************************
Function HandleError(boolTerminate,strAttachmentPath)
	If err.number >= 1000 And err.number < 2000 Then ' errors
		debug.WriteLine err.number & " " & err.Source & " " & err.Description
		ErrorMessage "<HTML><HEAD>" & sSharepointSourceDirectory & "</HEAD><BODY><p>" _
		 & "<span style=""color:red"">Error source</span>: " & err.Source _
		 & "<br><span style=""color:red"">Error code</span>: " & err.number _
		 & "<br><span style=""color:red"">Error description</span>: " & err.Description _
		 & "</p></BODY></HTML>","E;" & WScript.ScriptName & ";" & oDF.ToYearMonthDayWithDashes(Date),strAttachmentPath
		If boolTerminate Then WScript.Quit(err.number)
	ElseIf err.number >= 2000 And err.number < 3000 Then ' info/notifications
		debug.WriteLine err.number & " " & err.Source & " " & err.Description
		ErrorMessage "<HTML><HEAD>" & sSharepointSourceDirectory & "</HEAD><BODY><p>" _
		 & "<span style=""color:red"">Info source</span>: " & err.Source _
		 & "<br><span style=""color:red"">Info code</span>: " & err.number _
		 & "<br><span style=""color:red"">Info description</span>: " & err.Description _
		 & "</p></BODY></HTML>","I;" & WScript.ScriptName & ";" & oDF.ToYearMonthDayWithDashes(Date),strAttachmentPath
		If boolTerminate Then WScript.Quit(err.number)
	ElseIf err.number = 0 Then
	Else
		debug.WriteLine err.number & " " & err.Source & " " & err.Description
		ErrorMessage "<HTML><HEAD>" & sSharepointSourceDirectory & "</HEAD><BODY><p>" _
		 & "<span style=""color:red"">Error source</span>: " & err.Source _
		 & "<br><span style=""color:red"">Error code</span>: " & err.number _
		 & "<br><span style=""color:red"">Error description</span>: " & err.Description _
		 & "</p></BODY></HTML>","E;" & WScript.ScriptName & ";" & oDF.ToYearMonthDayWithDashes(Date),strAttachmentPath
		If boolTerminate Then WScript.Quit(err.number)
	End If
End Function 

Function UploadToSAP(ByRef ss,strSourceDir,strFile)

	If Not Right(strSourceDir,1) = "\" Then
		strSourceDir = strSourceDir & "\"
	End If 
	
	sl.KillPopups(ss)
	ss.findById("wnd[0]/tbar[0]/okcd").text="/NZTC_ZCVDOC47"
	ss.findById("wnd[0]").sendVKey 0
	sl.KillPopups(ss)
	ss.findById("wnd[0]/usr/ctxtDSNSET1").text= strSourceDir & strFile  		'path&filenamewithmax128characters
	ss.findById("wnd[0]/usr/txtW_SEPCHR").text=";"								'characterforseparation
	ss.findById("wnd[0]/usr/txtW_FEEDID").text=""								'FeederId(exit)
	ss.findById("wnd[0]/usr/txtW_FILEID").text=""								'FileId(exit)
	ss.findById("wnd[0]/usr/ctxtW_BUKRS").text= sCompanyCode					'CompanyCode
	
	If Right(strFile,7) = "_AR.csv" Then
		ss.findById("wnd[0]/usr/radW_TYPE2").select								'PostinginsubmoduleAR (TYPE2)
		ss.findById("wnd[0]/usr/ctxtW_BLART").text= sDocTypeAR					'Documenttype
		ss.findById("wnd[0]/usr/txtW_HTXT").text= sHeaderTextAR					'DocumentHeaderText
		ss.findById("wnd[0]/usr/txtP_SESS").text= sSessionAR		   			'Session name
	ElseIf Right(strFile,7) = "_AP.csv" Then
		ss.findById("wnd[0]/usr/radW_TYPE3").select								'PostinginsubmoduleAP (TYPE3)
		ss.findById("wnd[0]/usr/ctxtW_BLART").text= sDocTypeAP					'Documenttype
		ss.findById("wnd[0]/usr/txtW_HTXT").text= sHeaderTextAP					'DocumentHeaderText
		ss.findById("wnd[0]/usr/txtP_SESS").text= sSessionAP					'Session name
	ElseIf Right(strFile,7) = "APC.csv" Then
		ss.findById("wnd[0]/usr/radW_TYPE3").select								'PostinginsubmoduleAPC (TYPE3)
		ss.findById("wnd[0]/usr/ctxtW_BLART").text= sDocTypeAPC					'Documenttype
		ss.findById("wnd[0]/usr/txtW_HTXT").text= sHeaderTextAPC				'DocumentHeaderText
		ss.findById("wnd[0]/usr/txtP_SESS").text= sSessionAPC					'Session name
	End If 
	
	ss.findById("wnd[0]/usr/chkP_SUBM").selected=-1			'Automaticsubmitselected
	ss.findById("wnd[0]").sendVKey 8						'Execution F8
	sl.KillPopups(ss)
	
	Select Case ss.FindById(ss.ActiveWindow.Name & "/sbar").MessageType
	
		Case "E"
			debug.WriteLine "Error"
			ss.FindById(ss.ActiveWindow.Name).HardCopy "wnd",1
			oMAIL.SendMessageA "VIPSINT.UploadToSAP: " & ss.FindById(ss.ActiveWindow.Name & "/sbar/pane[0]").Text,"E;" & WScript.ScriptName & ";" & strSAPsys & ";" & oDF.ToYearMonthDayWithDashes(Date),strAttachment
			ss.findById("wnd[0]/tbar[0]/okcd").text="/nex"
			ss.findById("wnd[0]").sendVKey 0
			WScript.Quit
			
		Case "I"
			debug.WriteLine "Info"
			ss.FindById(ss.ActiveWindow.Name).HardCopy "wnd",1
			oMAIL.SendMessageA "VIPSINT.UploadToSAP: " & ss.FindById(ss.ActiveWindow.Name & "/sbar/pane[0]").Text,"I;" & WScript.ScriptName & ";" & strSAPsys & ";" & oDF.ToYearMonthDayWithDashes(Date),strAttachment

		Case "S"
			debug.WriteLine "Success"
			ss.FindById(ss.ActiveWindow.Name).HardCopy "wnd",1
			
		Case "W"
			debug.WriteLine "Warning"
			ss.FindById(ss.ActiveWindow.Name).HardCopy "wnd",1
			oMAIL.SendMessageA "VIPSINT.UploadToSAP: " & ss.FindById(ss.ActiveWindow.Name & "/sbar/pane[0]").Text,"W;" & WScript.ScriptName & ";" & strSAPsys & ";" & oDF.ToYearMonthDayWithDashes(Date),strAttachment
			
		Case "A"
			debug.WriteLine "Abort"
			ss.FindById(ss.ActiveWindow.Name).HardCopy "wnd",1
			oMAIL.SendMessageA "VIPSINT.UploadToSAP: " & ss.FindById(ss.ActiveWindow.Name & "/sbar/pane[0]").Text,"A;" & WScript.ScriptName & ";" & strSAPsys & ";" & oDF.ToYearMonthDayWithDashes(Date),strAttachment
			
		Case ""
			debug.WriteLine "No error message in the status bar. OK"
			
	End Select
		
	' If the code reaches this point, everything went ok
	UploadToSAP = True
	Exit Function 
	
End Function



Function ProcessSupplierCredit(strPath,ByRef arrFiles)
	Dim sCustomerType ' External or internal
	Dim arrFileNames(2)     ' A,B,C input file names
	Dim i,key,pos,field,line,errDescription,j,node
	Dim boolProcess : boolProcess = False ' Set to true if at least one entry found in the I file and therefore there is stuff to process
	Dim oStr : Set oStr = CreateObject("ADODB.Stream")
	Dim oCon : Set oCon = CreateObject("ADODB.Connection")
	Dim oCom : Set oCom = CreateObject("ADODB.Command")
	Dim oRst : Set oRst = CreateObject("ADODB.Recordset") ' Holds unique credit note number
	oRst.LockType = adLockOptimistic
	Dim oRstI : Set oRstI = CreateObject("ADODB.Recordset")
	Dim oRstJ : Set oRstJ = CreateObject("ADODB.Recordset")
	Dim oRx : Set oRx = New RegExp
	oRx.Pattern = "^\d+$"
	Dim oFile
	Dim glacc,netval
	Dim taxamount
	Dim lfDate,lfReference,lfTotalamount,lfParma,lfTradingpartner,lfProfitcenter,lfPaymentterm,lfTaxcode,lfSupplier,lfGlacc
	'#######################################
    'Configuration subtrees
    '#######################################
	Dim oSuppliers : Set oSuppliers = oConfigurationSubtree.SelectSingleNode("//Partners[@type='supplier']")
	Dim oTaxcodes : Set oTaxcodes = oConfigurationSubtree.SelectSingleNode("//TaxCodes")
	Dim oGLPC : Set oGLPC = oConfigurationSubtree.SelectSingleNode("//GLPCmatrix[@type='supplier']")
	
	
	For i = 0 To UBound(arrFiles)
		If LCase(Mid(arrFiles(i),Len(arrFiles(i)) - 4,1)) = "i" Then 
			arrFileNames(0) = arrFiles(i) 'D file at index 0
		ElseIf LCase(Mid(arrFiles(i),Len(arrFiles(i)) - 4,1)) = "j" Then
			arrFileNames(1) = arrFiles(i) 'E file at index 1
		End If 
	Next 
	
	With oCon
		.Provider = "Microsoft.ACE.OLEDB.16.0"
		.ConnectionString = "Data Source=" & strPath & ";Extended Properties=""Text;FMT=Delimited;"""
		.Open
	End With 
	
	' Select all distinct creditnote numbers
	oRst.Open "SELECT DISTINCT Creditnote FROM [" & arrFileNames(0) & "]", oCon
	Do While Not oRst.EOF
		If Not IsNull(oRst.Fields("Creditnote").Value) Then 
			If oRx.Test(oRst.Fields("Creditnote").Value) Then 
				boolProcess = True 
			End If 
		End If 
		oRst.MoveNext
	Loop
	
	' There are no records to be processed
	If Not boolProcess Then 
		ProcessSupplierCredit = False ' 0 or False. No entries found in the I file. NO output file will be created either
		Exit function
	End If 
	debug.WriteLine "boolProcess -> " & boolProcess & ". Entries found in the I file"
	
	'Set oFile = oFSO.OpenTextFile(sWorkingDirectory & "\SOURCE\" & sOutFilePrefixAPC & sOutFileName & sOutFileSuffixAPC,ForWriting,True) ' Create the output file
	
	oRst.MoveFirst ' Go back to the first record
	
	Do While Not oRst.EOF ' Loop through the distinct credit notes numbers
		If Not IsNull(oRst.Fields("Creditnote").Value) Then 
			If oRx.Test(oRst.Fields("Creditnote").Value) Then 
				'Select all I lines for each distinct credit note number
				oRstI.Open "SELECT Supplier, Dealer, SupplierType, Creditnote, InvoiceDate, InvoiceTotal FROM [" & arrFileNames(0) & "] WHERE Creditnote = '" & oRst.Fields("Creditnote").Value & "'", oCon		
				Do While Not oRstI.EOF ' Loop through all I lines that match distinct credit note number
					If Not IsNull(oRstI.Fields("Supplier").Value) Then 
						If oRx.Test(CStr(oRstI.Fields("Supplier").Value)) Then
							If oRstI.Fields("SupplierType").Value = 1 Then ' Only process type 1 suppliers
								If Not oFSO.FileExists(sWorkingDirectory & "\SOURCE\" & sOutFilePrefixAPC & sOutFileName & sOutFileSuffixAPC) Then
									oFSO.CreateTextFile sWorkingDirectory & "\SOURCE\" & sOutFilePrefixAPC & sOutFileName & sOutFileSuffixAPC,True 
									Set oFile = oFSO.OpenTextFile(sWorkingDirectory & "\SOURCE\" & sOutFilePrefixAPC & sOutFileName & sOutFileSuffixAPC,ForWriting,True) ' Create the output file
								End If 
								If Not oSuppliers.selectNodes("Partner[@number='" & oRstI.Fields("Supplier").Value & "' and @type='" & oRstI.Fields("SupplierType").Value & "']").length = 1 Then 
									'Missing partner
									'err.Raise 1001,"VIPSINT.ProcessSupplierCredit","Missing or duplicate partner " & oRstI.Fields("Supplier").Value & " detected "
									oMAIL.SendMessageText "VIPSINT.ProcessSupplierCredit" & vbCrLf & vbCrLf & "Missing or duplicate partner " & oRstI.Fields("Supplier").Value & " detected" ,"E;" & WScript.ScriptName & ";" & strSAPsys
									WScript.Quit
									Exit Function
								End If ' Error. Missing supplier
								
								'*************************
								'Process type 1 suppliers
								'Header lines
								'*************************
								lfTotalamount = lfTotalamount + CDbl(oRstI.Fields("InvoiceTotal").Value) ' Loop through the record set and calculate total amount of all invoices
								lfDate = oRstI.Fields("InvoiceDate").Value
								lfReference = oRst.Fields("Creditnote").Value
								lfSupplier = oRstI.Fields("Supplier").Value
								'Get PARMA,Tradingpartner & Paymentterm
								If Not oSuppliers.selectSingleNode("Partner[@number='" & oRstI.Fields("Supplier").Value & "' and @type='" & oRstI.Fields("SupplierType").Value & "']").childNodes.length = 3 Then
									'err.Raise 1002,"VIPSINT.ProcessSupplierCredit","Either PARMA, tradingpartner or paymentterm is missing for the partner " & oRstI.Fields("Supplier").Value
									oMAIL.SendMessageText "VIPSINT.ProcessSupplierCredit" & vbCrLf & vbCrLf & "Either PARMA, tradingpartner or paymentterm is missing for the partner " & oRstI.Fields("Supplier").Value,"E;" & WScript.ScriptName & ";" & strSAPsys
									WScript.Quit
									Exit Function
								End If 
								lfParma = oSuppliers.selectSingleNode("Partner[@number='" & oRstI.Fields("Supplier").Value & "' and @type='" & oRstI.Fields("SupplierType").Value & "']/parma").text
								lfTradingpartner = oSuppliers.selectSingleNode("Partner[@number='" & oRstI.Fields("Supplier").Value & "' and @type='" & oRstI.Fields("SupplierType").Value & "']/tradingpartner").text
								lfPaymentterm = oSuppliers.selectSingleNode("Partner[@number='" & oRstI.Fields("Supplier").Value & "' and @type='" & oRstI.Fields("SupplierType").Value & "']/paymentterm").text
								
								'Get Taxcode
								If Not oTaxcodes.selectNodes("TaxCode[@type='S']").length = 1 Then
									'err.Raise 1003,"VIPSINT.ProcessSupplierCredit","Missing or duplicate taxcodes of type 'S' detected"
									oMAIL.SendMessageText "VIPSINT.ProcessSupplierCredit" & vbCrLf & vbCrLf & "Missing or duplicate taxcodes of type 'S' detected","E;" & WScript.ScriptName & ";" & strSAPsys
									WScript.Quit
									Exit Function
								End If 
								lfTaxcode = oTaxcodes.selectSingleNode("TaxCode[@type='S']").text
		
								
							End If ' Type 1 supplier only
						End If ' Regex match rsti
					End If ' IsNull recordset value rsti
					oRstI.MoveNext
				Loop ' Loop through all I lines that match distinct credit note number
				oRstI.Close
				
				'After looping through oRstI recordset we should have all data needed to write out the header line and loop through J file
				line = lfDate & ";" & lfDate & ";" & sCurrency & ";" & lfReference & ";;" & Right("0000000000" & lfParma,10) & ";;" & FormatNumber(lfTotalamount,2,0,0,0) _
				     & ";" & lfTaxcode & ";0,00;;;;;;;;" & lfSupplier & ";;" & lfPaymentterm & ";;;;" _
				     & lfTradingpartner & ";;0,00;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;"
				     
				'reset total amount variable
				lfTotalamount = 0 
				'Write out the header line
				oFile.WriteLine line 
				
				'Query the J file
				If oRstJ.State = adStateOpen Then
					oRstJ.Close
				End If 
				oRstJ.Open "SELECT Supplier, Dealer, SupplierType, Creditnote, Netval, Invoice, Quantity, ProductGroup FROM [" & arrFileNames(1) & "] WHERE Creditnote = '" & oRst.Fields("Creditnote").Value & "'", oCON
				Do While Not oRstJ.EOF ' Loop through all J lines that match distinct credit note number 
					If Not IsNull(oRstJ.Fields("Supplier").Value) Then 
						If oRx.Test(oRstJ.Fields("Supplier").Value) Then 
							lfReference = oRst.Fields("Creditnote").Value
							lfTaxcode = oTaxcodes.selectSingleNode("TaxCode[@type='S']").text
							
							'Get GL account
							If Not oGLPC.selectNodes("GLPC[@number='" & oRstJ.Fields("Dealer").Value & "' and @productgroup='" & oRstJ.Fields("ProductGroup").Value & "']").length = 1 Then
								'err.Raise 1004,"VIPSINT.ProcessSupplierCredit","GLPC -> Missing or duplicate entries for the partner " & oRstJ.Fields("Dealer").Value & " productgroup " & oRstJ.Fields("ProductGroup").Value
								oMAIL.SendMessageText "VIPSINT.ProcessSupplierCredit" & vbCrLf & vbCrLf & "GLPC -> Missing or duplicate entries for the partner " & oRstJ.Fields("Dealer").Value & " productgroup " & oRstJ.Fields("ProductGroup").Value ,"E;" & WScript.ScriptName & ";" & strSAPsys
								WScript.Quit
								Exit Function
							End If 
							
							If Not oGLPC.selectSingleNode("GLPC[@number='" & oRstJ.Fields("Dealer").Value & "' and @productgroup='" & oRstJ.Fields("ProductGroup").Value & "']").childNodes.length = 2 Then
								'err.Raise 1005,"VIPSINT.ProcessSupplierCredit","GL/ProfitCenter -> Missing or duplicate entries for the partner " & oRstJ.Fields("Dealer").Value & " productgroup " & oRstJ.Fields("ProductGroup").Value
								oMAIL.SendMessageText "VIPSINT.ProcessSupplierCredit" & vbCrLf & vbCrLf & "GL/ProfitCenter -> Missing or duplicate entries for the partner " & oRstJ.Fields("Dealer").Value & " productgroup " & oRstJ.Fields("ProductGroup").Value ,"E;" & WScript.ScriptName & ";" & strSAPsys
								WScript.Quit
								Exit Function
							End If 
							
							lfGlacc = oGLPC.selectSingleNode("GLPC[@number='" & oRstJ.Fields("Dealer").Value & "' and @productgroup='" & oRstJ.Fields("ProductGroup").Value & "']/GL").text
							lfProfitcenter = oGLPC.selectSingleNode("GLPC[@number='" & oRstJ.Fields("Dealer").Value & "' and @productgroup='" & oRstJ.Fields("ProductGroup").Value & "']/ProfitCenter").text
						
							line = lfDate & ";" & lfDate & ";" & sCurrency & ";" & lfReference & ";" & Right("0000000000" & lfGlacc,10) & ";;;" _
								 & FormatNumber(CDbl(oRstJ.Fields("Netval").Value) * CDbl(oRstJ.Fields("Quantity").Value),2,0,0,0) & "-;" & lfTaxcode & ";0,00;;" _
								 & Right("0000000000" & lfProfitcenter,10) & ";;;;;;" & oRstJ.Fields("Dealer").Value & "/" & oRstJ.Fields("Invoice").Value _
								 & ";;;;;;;;0,00;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;"
								 
							oFile.WriteLine line
						End If ' Line contains valid data not garbage
					End If ' Line contains valid data not garbage
					oRstJ.MoveNext
				Loop ' Loop through all J lines that match distinct credit note number 
			End If ' Regex match rst 
		End If ' IsNull recordset value rst
		oRst.MoveNext
	Loop  ' Loop through the distinct credit notes numbers
	
	oCon.Close
	
	If IsObject(oFile) Then
		oFile.Close
	Else 
		ProcessSupplierCredit = False ' Entries found, but no external customer exists within the file. Return False
		Exit Function 
	End If 
	
	ProcessSupplierCredit = True
End Function 
	


Function ProcessSupplier(strPath,ByRef arrFiles)
	Dim sCustomerType ' External or internal
	Dim arrFileNames(2)     ' A,B,C input file names
	Dim i,key,pos,field,line,errDescription,j,node
	Dim boolProcess : boolProcess = False ' Set to true if at least one entry foun in the D file and therefore there is stuff to process
	Dim oStr : Set oStr = CreateObject("ADODB.Stream")
	Dim oCon : Set oCon = CreateObject("ADODB.Connection")
	Dim oCom : Set oCom = CreateObject("ADODB.Command")
	Dim oRstD : Set oRstD = CreateObject("ADODB.Recordset")
	Dim oRstE : Set oRstE = CreateObject("ADODB.Recordset")
	Dim oPartners : Set oPartners = CreateObject("Scripting.Dictionary")
	Dim oFile
	Dim oRx : Set oRx = New RegExp
	oRx.Pattern = "^\d+$"
	Dim parma,tradingpartner,profitcenter,glacc,netval,paymentterm
	Dim taxcode,taxamount
	'#######################################
    'Configuration subtrees
    '#######################################
	Dim oSuppliers : Set oSuppliers = oConfigurationSubtree.SelectSingleNode("//Partners[@type='supplier']")
	Dim oTaxcodes : Set oTaxcodes = oConfigurationSubtree.SelectSingleNode("//TaxCodes")
	Dim oGLPC : Set oGLPC = oConfigurationSubtree.SelectSingleNode("//GLPCmatrix[@type='supplier']")
	
	
	For i = 0 To UBound(arrFiles)
		If LCase(Mid(arrFiles(i),Len(arrFiles(i)) - 4,1)) = "d" Then 
			arrFileNames(0) = arrFiles(i) 'D file at index 0
		ElseIf LCase(Mid(arrFiles(i),Len(arrFiles(i)) - 4,1)) = "e" Then
			arrFileNames(1) = arrFiles(i) 'E file at index 1
		End If 
	Next 			
	
		
	With oCon
		.Provider = "Microsoft.ACE.OLEDB.16.0"
		.ConnectionString = "Data Source=" & strPath & ";Extended Properties=""Text;FMT=Delimited;"""
		.Open
	End With 
	
	oRstD.Open "SELECT Partner, PartnerType, Invoice, InvoiceDate, InvoiceTotal FROM [" & arrFileNames(0) & "]", oCON
	
	Do While Not oRstD.EOF
		If Not IsNull(oRstD.Fields("Partner").Value) Then 
			If oRx.Test(oRstD.Fields("Partner").Value) Then 
				boolProcess = True 
			End If 
		End If 
		oRstD.MoveNext
	Loop
	
	' There are no records to be processed
	If Not boolProcess Then 
		ProcessSupplier = False ' 0 or False. No entries found in the D file. NO output file will be created either
		Exit function
	End If 
	debug.WriteLine "boolProcess -> " & boolProcess & ". Entries found in the D file"
	'Set oFile = oFSO.OpenTextFile(sWorkingDirectory & "\SOURCE\" & sOutFilePrefixAP & sOutFileName & sOutFileSuffixAP,ForWriting,True)
	
	oRstD.MoveFirst

	Do While Not oRstD.EOF
		If Not IsNull(oRstD.Fields("Partner").Value) Then 
			If oRx.Test(CStr(oRstD.Fields("Partner").Value)) Then
				If oRstD.Fields("PartnerType").Value = 1 Then ' Only process type 1 suppliers
					If Not oFSO.FileExists(sWorkingDirectory & "\SOURCE\" & sOutFilePrefixAP & sOutFileName & sOutFileSuffixAP) Then
						Set oFile = oFSO.OpenTextFile(sWorkingDirectory & "\SOURCE\" & sOutFilePrefixAP & sOutFileName & sOutFileSuffixAP,ForWriting,True)
					End If 
					If Not oSuppliers.selectNodes("Partner[@number='" & oRstD.Fields("Partner").Value & "' and @type='" & oRstD.Fields("PartnerType").Value & "']").length = 1 Then 
						'Missing partner
						'err.Raise 1001,"VIPSINT.ProcessSupplier","Missing or duplicate partner " & oRstD.Fields("Partner").Value & " detected "
						oMAIL.SendMessage "VIPSINT.ProcessSupplier" & vbCrLf & vbCrLf & "Missing or duplicate partner " & oRstD.Fields("Partner").Value & " detected " ,"E;" & WScript.ScriptName & ";" & strSAPsys
						WScript.Quit
					End If ' Error. Missing supplier
					
					'*************************
					'Process type 1 suppliers
					'Header lines
					'*************************
					
					line = oRstD.Fields("InvoiceDate").Value & ";" & oRstD.Fields("InvoiceDate").Value & ";" & sCurrency & ";" & oRstD.Fields("Invoice").Value & ";;"
					
					'Get PARMA
					If Not oSuppliers.selectSingleNode("Partner[@number='" & oRstD.Fields("Partner").Value & "' and @type='" & oRstD.Fields("PartnerType").Value & "']").childNodes.length = 3 Then
						'err.Raise 1002,"VIPSINT.ProcessSupplier","Either PARMA, tradingpartner or paymentterm is missing for the partner " & oRstD.Fields("Partner").Value
						oMAIL.SendMessage "VIPSINT.ProcessSupplier" & vbCrLf & vbCrLf & "Either PARMA, tradingpartner or paymentterm is missing for the partner " & oRstD.Fields("Partner").Value,"E;" & WScript.ScriptName & ";" & strSAPsys
						WScript.Quit
						Exit Function
					End If 
					
					parma = oSuppliers.selectSingleNode("Partner[@number='" & oRstD.Fields("Partner").Value & "' and @type='" & oRstD.Fields("PartnerType").Value & "']/parma").text
					tradingpartner = oSuppliers.selectSingleNode("Partner[@number='" & oRstD.Fields("Partner").Value & "' and @type='" & oRstD.Fields("PartnerType").Value & "']/tradingpartner").text
					paymentterm = oSuppliers.selectSingleNode("Partner[@number='" & oRstD.Fields("Partner").Value & "' and @type='" & oRstD.Fields("PartnerType").Value & "']/paymentterm").text
					
					line = line & Right("0000000000" & parma,10) & ";;" & oRstD.Fields("InvoiceTotal").Value & "-;"
					
					'Get Taxcode
					If Not oTaxcodes.selectNodes("TaxCode[@type='S']").length = 1 Then
						'err.Raise 1003,"VIPSINT.ProcessSupplier","Missing or duplicate taxcodes of type 'S' detected"
						oMAIL.SendMessage "VIPSINT.ProcessSupplier" & vbCrLf & vbCrLf & "Missing or duplicate taxcodes of type 'S' detected","E;" & WScript.ScriptName & ";" & strSAPsys
						WScript.Quit
						Exit Function
					End If 
					
					taxcode = oTaxcodes.selectSingleNode("TaxCode[@type='S']").text
					
					line = line & taxcode & ";0,00;;;;;;;;" & oRstD.Fields("Partner").Value & ";;" & paymentterm & ";;;;" & tradingpartner & ";;0,00;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;"
					
					oFile.WriteLine line
					
					'*************************
					'Process type 1 suppliers
					'GL lines
					'*************************
					oRstE.Open "SELECT Partner, DealerNumber, ProductGroup, NetValue, Quantity, Invoice, InvoiceDate, InvoiceTotal FROM [" & arrFileNames(1) & "] WHERE Partner ='" & oRstD.Fields("Partner").Value & "' AND Invoice = '" & oRstD.Fields("Invoice").Value & "'", oCon
					Do While Not oRstE.EOF
						line = oRstD.Fields("InvoiceDate").Value & ";" & oRstD.Fields("InvoiceDate").Value & ";" & sCurrency & ";" & oRstD.Fields("Invoice").Value & ";"
						
						'Get GL account
						If Not oGLPC.selectNodes("GLPC[@number='" & oRstE.Fields("DealerNumber").Value & "' and @productgroup='" & oRstE.Fields("ProductGroup").Value & "']").length = 1 Then
							'err.Raise 1004,"VIPSINT.ProcessSupplier","GLPC -> Missing or duplicate entries for the partner " & oRstE.Fields("DealerNumber").Value & " productgroup " & oRstE.Fields("ProductGroup").Value
							oMAIL.SendMessage "VIPSINT.ProcessSupplier" & vbCrLf & vbCrLf & "GLPC -> Missing or duplicate entries for the partner " & oRstE.Fields("DealerNumber").Value & " productgroup " & oRstE.Fields("ProductGroup").Value,"E;" & WScript.ScriptName & ";" & strSAPsys
							WScript.Quit
							Exit Function
						End If 
						
						If Not oGLPC.selectSingleNode("GLPC[@number='" & oRstE.Fields("DealerNumber").Value & "' and @productgroup='" & oRstE.Fields("ProductGroup").Value & "']").childNodes.length = 2 Then
							'err.Raise 1005,"VIPSINT.ProcessSupplier","GL/ProfitCenter -> Missing or duplicate entries for the partner " & oRstE.Fields("DealerNumber").Value & " productgroup " & oRstE.Fields("ProductGroup").Value
							oMAIL.SendMessage "VIPSINT.ProcessSupplier" & vbCrLf & vbCrLf & "GL/ProfitCenter -> Missing or duplicate entries for the partner " & oRstE.Fields("DealerNumber").Value & " productgroup " & oRstE.Fields("ProductGroup").Value,"E;" & WScript.ScriptName & ";" & strSAPsys
							WScript.Quit
							Exit Function
						End If 
						
						glacc = oGLPC.selectSingleNode("GLPC[@number='" & oRstE.Fields("DealerNumber").Value & "' and @productgroup='" & oRstE.Fields("ProductGroup").Value & "']/GL").text
						profitcenter = oGLPC.selectSingleNode("GLPC[@number='" & oRstE.Fields("DealerNumber").Value & "' and @productgroup='" & oRstE.Fields("ProductGroup").Value & "']/ProfitCenter").text
						netval = FormatNumber(CDbl(oRstE.Fields("NetValue").Value) * CDbl(oRstE.Fields("Quantity").Value),2,0,0,0) & ";"
						
						line = line & Right("000000000" & glacc,10) & ";;;" & netval & taxcode & ";0,00;;" & Right("0000000000" & profitcenter,10) & ";;;;;;" & oRstE.Fields("DealerNumber").Value & ";;;;;;;;" & netval & ";;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;"
						
						oFile.WriteLine line
						
						oRstE.MoveNext
					Loop
					oRstE.Close 
				End If ' Type 1 supplier only
			End If ' Regex match
		End If ' IsNull recordset value
		oRstD.MoveNext
	Loop 
	
	oCon.Close
	
	If IsObject(oFile) Then
		oFile.Close
	Else 
		ProcessSupplier = False ' Entries found, but no external customer exists within the file. Return False
		Exit Function 
	End If 
	
	ProcessSupplier = True
End Function 


Function ProcessCustomer(strPath,ByRef arrFiles)
	Dim sCustomerType ' External or internal
	Dim arrFileNames(3)     ' A,B,C input file names
	Dim i,key,pos,field,line,errDescription,j,node
	Dim boolProcess : boolProcess = False ' Set to true if at least one entry foun in the A file and therefore there is stuff to process
	Dim oStr : Set oStr = CreateObject("ADODB.Stream")
	Dim oCon : Set oCon = CreateObject("ADODB.Connection")
	Dim oCom : Set oCom = CreateObject("ADODB.Command")
	Dim oRstA : Set oRstA = CreateObject("ADODB.Recordset")
	Dim oRstB : Set oRstB = CreateObject("ADODB.Recordset")
	Dim oRstC : Set oRstC = CreateObject("ADODB.Recordset")
	Dim oPartners : Set oPartners = CreateObject("Scripting.Dictionary")
	Dim oFile
	Dim oRx : Set oRx = New RegExp
	oRx.Pattern = "^\d+$"
	Dim parma,tradingpartner,profitcenter,glacc,netval
	Dim taxcode,taxamount,bentries
	'#######################################
    'Configuration subtrees
    '#######################################
	Dim oCustomers : Set oCustomers = oConfigurationSubtree.SelectSingleNode("//Partners[@type='customer']")
	Dim oTaxcodes : Set oTaxcodes = oConfigurationSubtree.SelectSingleNode("//TaxCodes")
	Dim oGLPC : Set oGLPC = oConfigurationSubtree.SelectSingleNode("//GLPCmatrix[@type='customer']")
	
	
	For i = 0 To UBound(arrFiles)
		If LCase(Mid(arrFiles(i),Len(arrFiles(i)) - 4,1)) = "a" Then 
			arrFileNames(0) = arrFiles(i) 'A file at index 0
		ElseIf LCase(Mid(arrFiles(i),Len(arrFiles(i)) - 4,1)) = "b" Then
			arrFileNames(1) = arrFiles(i) 'B file at index 1
		ElseIf LCase(Mid(arrFiles(i),Len(arrFiles(i)) - 4,1)) = "c" Then
			arrFileNames(2) = arrFiles(i) 'C file at index 2
		End If 
	Next 			
	
	
	With oCon
		.Provider = "Microsoft.ACE.OLEDB.16.0"
		.ConnectionString = "Data Source=" & strPath & ";Extended Properties=""Text;FMT=Delimited;"""
		.Open
	End With 
	
	oRstA.Open "SELECT Partner, Invoice, InvoiceType, InvoiceDate, InvoiceTotal, Payterm FROM [" & arrFileNames(0) & "]", oCON
	
	Do While Not oRstA.EOF
		If Not IsNull(oRstA.Fields("Partner").Value) Then 
			If oRx.Test(oRstA.Fields("Partner").Value) Then 
				boolProcess = True 
			End If 
		End If 
		oRstA.MoveNext
	Loop
	
	' There are no records to be processed
	If Not boolProcess Then 
		ProcessCustomer = False ' 0 or False. No entries found in the I file. NO output file will be created either
		Exit function
	End If 
	debug.WriteLine "boolProcess -> " & boolProcess & ". Entries found in the A file"
	'Set oFile = oFSO.OpenTextFile(sWorkingDirectory & "\SOURCE\" & sOutFilePrefixAR & sOutFileName & sOutFileSuffixAR,ForWriting,True)
	
	oRstA.MoveFirst
	
	Do While Not oRstA.EOF
		If Not IsNull(oRstA.Fields("Partner").Value) Then 
			If oRx.Test(CStr(oRstA.Fields("Partner").Value)) Then
				If Not oCustomers.selectNodes("Partner[@number='" & oRstA.Fields("Partner").Value & "']").length = 1 Then 
					'Missing partner
					'err.Raise 1001,"VIPSINT.ProcessCustomer","Missing or duplicate partner " & oRstD.Fields("Partner").Value & " detected "
					oMAIL.SendMessageText "VIPSINT.ProcessCustomer" & vbCrLf & vbCrLf & "Missing or duplicate partner " & oRstA.Fields("Partner").Value & " detected ","E;" & WScript.ScriptName & ";" & strSAPsys
					WScript.Quit
				End If 
				
				'Process only external customer
				If oCustomers.selectSingleNode("Partner[@number='" & oRstA.Fields("Partner").Value & "']").attributes.getNamedItem("type").text = "E" Then
					'Only create output file for writing if there is external customer
					If Not oFSO.FileExists(sWorkingDirectory & "\SOURCE\" & sOutFilePrefixAR & sOutFileName & sOutFileSuffixAR) Then
						oFSO.CreateTextFile sWorkingDirectory & "\SOURCE\" & sOutFilePrefixAR & sOutFileName & sOutFileSuffixAR,True
						Set oFile = oFSO.OpenTextFile(sWorkingDirectory & "\SOURCE\" & sOutFilePrefixAR & sOutFileName & sOutFileSuffixAR,ForWriting,True)
					End If 
					If oCustomers.selectSingleNode("Partner[@number='" & oRstA.Fields("Partner").Value & "']").childNodes.length < 2 Then
						'Missing either parma or tradingpartner
						'err.Raise 1002,"VIPSINT.ProcessCustomer","Partner " & oRstA.Fields("Partner").Value & " is missing either parma or tradingpartner"
						oMAIL.SendMessageText "VIPSINT.ProcessCustomer" & vbCrLf & vbCrLf & "Partner " & oRstA.Fields("Partner").Value & " is missing either parma or tradingpartner","E;" & WScript.ScriptName & ";" & strSAPsys
						WScript.Quit
					End If
					'*****************************************
					'Obtain data from B file and configuration
					'*****************************************
					parma = oCustomers.selectSingleNode("Partner[@number='" & oRstA.Fields("Partner").Value & "']/parma").text
					tradingpartner = oCustomers.selectSingleNode("Partner[@number='" & oRstA.Fields("Partner").Value & "']/tradingpartner").text
					line = oRstA.Fields("InvoiceDate").Value & ";" & oRstA.Fields("InvoiceDate").Value & ";" & sCurrency & oRstA.Fields("Invoice").Value _
					       & sCustRefSuffix & ";" & ";" & Right("0000000000" & parma,10) & ";"
					
					'Ivnoice/Creditnote total  
					If oRstA.Fields("InvoiceType").Value = 1 Or oRstA.Fields("InvoiceType").Value = 2 Then
						line = line & oRstA.Fields("InvoiceTotal").Value & ";"
					Else
						line = line & oRstA.Fields("InvoiceTotal").Value & "-;"
					End If 
					
					'Get taxcode from B
					oRstB.Open "SELECT Taxcode, Taxamount FROM [" & arrFileNames(1) & "] WHERE Partner = '" & oRsta.Fields("Partner").Value & "' AND Invoice = '" & oRstA.Fields("Invoice").Value & "'", oCon
					bentries = 0
					Do While Not oRstB.EOF
						oRstB.MoveNext
						bentries = bentries + 1
					Loop 
				
					'Check if there are valid entries in the B file for Parnter + Invoice combination
					If bentries = 0 Or bentries > 1 Then
						'err.Raise 1003,"VIPSINT.ProcessCustomer","Missing or duplicate entries in the file " & arrFileNames(1) & " for the partner " & oRstA.Fields("Partner").Value & " invoice " & oRstA.Fields("Invoice").Value
						oMAIL.SendMessageText "VIPSINT.ProcessCustomer" & vbCrLf & vbCrLf & "Missing or duplicate entries in the file " & arrFileNames(1) & " for the partner " & oRstA.Fields("Partner").Value & " invoice " & oRstA.Fields("Invoice").Value,"E;" & WScript.ScriptName & ";" & strSAPsys
						WScript.Quit
						Exit Function
					Else
						oRstB.MoveFirst
						taxcode = oRstB.Fields("Taxcode").Value
						taxamount = oRstB.Fields("Taxamount").Value
					End If 
					
					oRstB.Close
					
					'Get taxcode from configuration file
					If Not oTaxcodes.selectNodes("TaxCode[@type='C' and @VIPS_Code='" & taxcode & "']").length = 1 Then
						'err.Raise 1004,"VIPSINT.ProcessCustomer","TaxCode -> Missing or duplicate entries for the partner " & oRstA.Fields("Partner").Value
						oMAIL.SendMessageText "VIPSINT.ProcessCustomer" & vbCrLf & vbCrLf & "TaxCode -> Missing or duplicate entries for the partner " & oRstA.Fields("Partner").Value,"E;" & WScript.ScriptName & ";" & strSAPsys
						WScript.Quit
						Exit Function
					Else
						taxcode = oTaxcodes.selectSingleNode("TaxCode[@type='C' and @VIPS_Code='" & taxcode & "']").text
					End If 
					
					'Write data to file
					line = oRstA.Fields("InvoiceDate").Value & ";" & oRstA.Fields("InvoiceDate").Value & ";" & sCurrency & ";" _
						   & oRstA.Fields("Invoice").Value & sCustRefSuffix & ";;" & parma & ";;"
						   
				    ' Invoice total value based on invoice type 
					If oRstA.Fields("InvoiceType") = "1" Or oRstA.Fields("InvoiceType").Value = "2" Then
						line = line & oRstA.Fields("InvoiceTotal").Value & ";" ' Invoice -> Positive
					Else
						line = line & oRstA.Fields("InvoiceTotal").Value & "-;" ' Credit note -> Negative
					End If 
					
					line = line & taxcode & ";" & taxamount & ";;;;;;;;" & oRstA.Fields("Partner").Value & ";;C0" & oRstA.Fields("Payterm").Value _
					      & ";;;;" & tradingpartner & ";;0,00;0,00;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;"
					
					'********************************
					' Write out the header line
					'********************************   
					oFile.WriteLine line
					
					'********************************
					' Continue with C file
					'********************************
					oRstC.Open "SELECT Partner, Invoice, NetValue, Quantity, Taxcode, ProductGroup  FROM [" & arrFileNames(2) & "] WHERE Partner = '" & oRsta.Fields("Partner").Value & "' AND Invoice = '" & oRstA.Fields("Invoice").Value & "'", oCon
					Do While Not oRstC.EOF
						line = oRstA.Fields("InvoiceDate").Value & ";" & oRstA.Fields("InvoiceDate").Value & ";" & sCurrency & ";" & oRstC.Fields("Invoice").Value & sCustRefSuffix & ";"
						
						' Get GL Account from configuration
						If Not oGLPC.selectNodes("GLPC[@number='" & oRstC.Fields("Partner").Value & "' and @productgroup='" & oRstC.Fields("ProductGroup").Value & "']").length = 1 Then
							'err.Raise 1005, "VIPSINT.ProcessCustomer","GLPC -> Missing or duplicate entries for the partner " & oRstC.Fields("Partner").Value & " productgroup " & oRstC.Fields("ProductGroup").Value
							oMAIL.SendMessageText "VIPSINT.ProcessCustomer" & vbCrLf & vbCrLf & "GLPC -> Missing or duplicate entries for the partner " & oRstC.Fields("Partner").Value & " productgroup " & oRstC.Fields("ProductGroup").Value,"E;" & WScript.ScriptName & ";" & strSAPsys
							WScript.Quit
							Exit Function
						ElseIf Not oGLPC.selectSingleNode("GLPC[@number='" & oRstC.Fields("Partner").Value & "' and @productgroup='" & oRstC.Fields("ProductGroup").Value & "']").childNodes.length = 2 Then
							'err.Raise 1006, "VIPSINT.ProcessCustomer","GL/ProfitCenter -> Missing or duplicate entries for the partner " & oRstC.Fields("Partner").Value & " productgroup " & oRstC.Fields("ProductGroup").Value
							oMAIL.SendMessageText "VIPSINT.ProcessCustomer" & vbCrLf & vbCrLf & "GL/ProfitCenter -> Missing or duplicate entries for the partner " & oRstC.Fields("Partner").Value & " productgroup " & oRstC.Fields("ProductGroup").Value,"E;" & WScript.ScriptName & ";" & strSAPsys
							WScript.Quit
							Exit Function
						Else 
							glacc = oGLPC.selectSingleNode("GLPC[@number='" & oRstC.Fields("Partner").Value & "' and @productgroup='" & oRstC.Fields("ProductGroup").Value & "']/GL").text
							profitcenter = oGLPC.selectSingleNode("GLPC[@number='" & oRstC.Fields("Partner").Value & "' and @productgroup='" & oRstC.Fields("ProductGroup").Value & "']/ProfitCenter").text
						End If 
						
						line = line & glacc & ";;;"
						
						' NetValue * Quantity
						If oRstA.Fields("InvoiceType").Value = 1 Or oRstA.Fields("InvoiceType").Value = 2 Then
							netval = FormatNumber(CDbl(oRstC.Fields("NetValue").Value) * CDbl(oRstC.Fields("Quantity").Value),2,0,0,0) & "-;"
						Else
							netval = FormatNumber(CDbl(oRstC.Fields("NetValue").Value) * CDbl(oRstC.Fields("Quantity").Value),2,0,0,0) & ";"
						End If 
						
						taxcode = oTaxcodes.selectSingleNode("TaxCode[@type='C' and @VIPS_Code='" & oRstC.Fields("Taxcode").Value & "']").text
						
						line = line & netval & taxcode & ";0,00;;" & Right("0000000000" & profitcenter,10) & ";;;;;;" & oRstC.Fields("Partner").Value & "_" & oRstC.Fields("Invoice").Value & ";;;;;;;;" & netval & "0,00;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;"
						
						oFile.WriteLine line 
						oRstC.MoveNext
					Loop ' oRstC
					oRstC.Close
						
				End If ' External customer section
			Else ' Not valid parnter number. Either empty line or any other line w/o data
				' Garbage line
			End If
			
			
		End If   
		oRstA.MoveNext
	Loop 
	
	oCon.Close
	
	If IsObject(oFile) Then
		oFile.Close
	Else 
		ProcessCustomer = False ' Entries found, but no external customer exists within the file. Return False
		Exit Function 
	End If 
	
	ProcessCustomer = True
	
	
End Function 


Sub Download()
	oSP.Init sHost,sDomain,sUsername,sSecret
	retval = oSP.DownloadFilesA(sSharepointSourceDirectory,sWorkingDirectory & "\SOURCE",arrFilesToDownload)
	If retval < 0 Then
		debug.WriteLine "Error occured"
		debug.WriteLine "Error source: " & oSP.LastErrorSource
		debug.WriteLine "Error code: " & oSP.LastErrorCode
		debug.WriteLine "Error description: " & oSP.LastErrorDesc
		ErrorMessage "<HTML><HEAD>" & sSharepointSourceDirectory & "</HEAD><BODY><p>" _
		 & "<span style=""color:red"">Error source</span>: " & oSP.LastErrorSource _
		 & "<br><span style=""color:red"">Error code</span>: " & oSP.LastErrorCode _
		 & "<br><span style=""color:red"">Error description</span>: " & oSP.LastErrorDesc _
		 & "</p></BODY></HTML>","E;" & WScript.ScriptName & ";" & strSAPsys & ";" & oDF.ToYearMonthDayWithDashes(Date),Null
		WScript.Quit(oSP.LastErrorCode)
	End If
End Sub 


'---------------------------
'ErrorMessage subroutine
'---------------------------
Sub ErrorMessage(strBody,strSubject,strAttachment)
	If Not IsNull(strAttachment) Then
		oMAIL.SendMessage strBody,strSubject
	Else 
		oMAIL.SendMessage strBody,strSubject
	End If 
End Sub 

'------------------
' Mailer class
'------------------
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
		
		strAdmins = ""
		strUserName = oNET.UserName
		strComputerName = oNET.ComputerName
		
	End Sub 
	
	Private Sub Class_Terminate
	
	End Sub 
	
	Public Function SendMessageA(strMessage,strSubject,strAttach)
		
			If strAdmins = "" Then
				Exit Function
			End If 
			
    		Dim oUser : Set oUser = GetObject("LDAP://" & oSysInfo.UserName)
	
			With oEmail 
				.From = oUser.Mail
				.To = strAdmins
				.Subject = strSubject
				.Configuration.Fields.Item _
				("http://schemas.microsoft.com/cdo/configuration/sendusing") = 2
				.Configuration.Fields.Item _
				("http://schemas.microsoft.com/cdo/configuration/smtpserver") = _
	    		"mailgot.it.volvo.net" 
				.Configuration.Fields.Item _
	    		("http://schemas.microsoft.com/cdo/configuration/smtpserverport") = 25
	    		'.TextBody = strMessage
	    		.HTMLBody = strMessage
	    		.AddAttachment strAttach
	    		.Configuration.Fields.Item("urn:schemas:mailheader:X-MSMail-Priority") = "High"
	    		.Configuration.Fields.Item("urn:schemas:httpmail:importance") = 2
				.Configuration.Fields.Item("urn:schemas:mailheader:X-Priority") = 2
				.Configuration.Fields.Update
				.Send
			End With 
			
	End Function 
    
	
	Public Function SendMessage(strMessage,strSubject)
	
		If strAdmins = "" Then
			Exit Function
		End If 
			
		Dim oUser : Set oUser = GetObject("LDAP://" & oSysInfo.UserName)
	
		With oEmail 
			.From = oUser.Mail
			.To = strAdmins
			.Subject = strSubject
			.Configuration.Fields.Item _
			("http://schemas.microsoft.com/cdo/configuration/sendusing") = 2
			.Configuration.Fields.Item _
			("http://schemas.microsoft.com/cdo/configuration/smtpserver") = _
    		"mailgot.it.volvo.net" 
			.Configuration.Fields.Item _
    		("http://schemas.microsoft.com/cdo/configuration/smtpserverport") = 25
    		'.TextBody = strMessage
    		.HTMLBody = strMessage
    		.Configuration.Fields.Item("urn:schemas:mailheader:X-MSMail-Priority") = "High"
    		.Configuration.Fields.Item("urn:schemas:httpmail:importance") = 2
			.Configuration.Fields.Item("urn:schemas:mailheader:X-Priority") = 2
			.Configuration.Fields.Update
			.Send
		End With 
			
	
	End Function
	
	Public Function SendMessageText(strMessage,strSubject)
	
		If strAdmins = "" Then
			Exit Function
		End If
		 
		Dim oUser : Set oUser = GetObject("LDAP://" & oSysInfo.UserName)
	
		With oEmail 
			.From = oUser.Mail
			.To = strAdmins
			.Subject = strSubject
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
			
	
	End Function
	
	Public Property Let AddAdmin(strEmailAddress)
		strAdmins = strAdmins & strEmailAddress & ";"
	End Property 
	
	Public Property Get GetAdmins
		GetAdmins = Left(strAdmins,Len(strAdmins) - 1)
	End Property 
	
End Class 


'---------------------
' DateFormatter class
'---------------------
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


'------------------
' StopWatch class
'------------------
Class StopWatch

	Private start
	Private finish
	
	Private Sub Class_Initialize
	
		start = 0
		finish = 0
		
	End Sub
	
	Private Sub Class_Terminate
	
	End Sub 
	

	Public Function Activate
		start = Hour(Now()) * 3600 + Minute(Now()) * 60 + Second(Now())
	End Function
	
	Public Function Deactivate
		finish = Hour(Now()) * 3600 + Minute(Now()) * 60 + Second(Now())
	End Function 
	
	
	Public Property Get Duration ' Seconds
		Duration = finish - start
	End Property 
	
	
End Class

'********************** SP.vbs ******************************
' Class for manipulating Sharepoint online using REST API
' Initialize object with your site URL, your organization domain, Client ID and Client Secret
' Generate your Client ID and Client Secret when you register your app 
' Register your app here: https://{your organization}.sharepoint.com/sites/{your site}/_layouts/15/appregnew.aspx





Class SP
	
	Private oXML
	Private strAuthUrlPart1
	Private strAuthUrlPart2
	Private vti_bin_clientURL
	Private oHTTP
	Private oFSO
	Private strClientID
	Private strSecurityToken
	Private strClientSecret
	Private strFormDigestValue
	Private strTenantID
	Private strResourceID
	Private strURLbody
	Private strSiteURL
	Private strDomain
	Private numHTTPstatus
	Private strSite
	Private errDescription
	Private errNumber
	Private errSource
	
	Private Sub Class_Initialize
		
		errDescription = ""
		errNumber = 0
		numHTTPstatus = 0
		strAuthUrlPart1 = "https://accounts.accesscontrol.windows.net/"
		strAuthUrlPart2 = "/tokens/OAuth/2"
		Set oFSO = CreateObject("Scripting.FileSystemObject")
		Set oHTTP = CreateObject("MSXML2.ServerXMLHTTP.3.0")
		Set oXML = CreateObject("MSXML2.DOMDocument")
	End Sub 
	
	Public Function Init(sSiteUrl,sDomain,sClientID,sClientSecret)
	
		Dim parts
		If Right(sSiteUrl,1) = "/" Then
			strSiteURL = sSiteUrl
		Else
			strSiteURL = sSiteUrl & "/"
		End If
		
		parts = Split(strSiteURL,"/")
		
		strSite = parts(UBound(parts) - 1) 
		
		If Left(sDomain,1) = "/" Then
			sDomain = Right(sDomain,Len(sDomain) - 1)
		End If
		If Right(sDomain,1) = "/" Then
			sDomain = Left(sDomain,Len(sDomain) - 1)
		End If 
		strDomain = sDomain
		strClientID = sClientID
		strClientSecret = sClientSecret
		
		If Right(sSiteUrl,1) = "/" Then
			vti_bin_clientURL = sSiteUrl & "_vti_bin/client.svc"
		Else
			vti_bin_clientURL = sSiteUrl & "/_vti_bin/client.svc"
		End If 
		
		GetTenantID      ' Obtain the Tenant/Realm ID
		GetSecurityToken ' Obtain the Security Token
		GetXDigestValue  ' Obtain the form digest value
	
	End Function 
		
	'********************** P R I V A T E   F U N C T I O N S ************************
	Private Function GetTenantID()
	
		Dim part,parts,header
		oHTTP.open "GET",vti_bin_clientURL,False
		oHTTP.setRequestHeader "Authorization","Bearer"
		oHTTP.send
	
		parts = Split(oHTTP.getResponseHeader("WWW-Authenticate"),",")
	
		For Each part In parts 
	
			If InStr(part,"Bearer realm") > 0 Then
				header = Split(part,"=")
				strTenantID = header(1)
				strTenantID = Mid(strTenantID,2,Len(strTenantID) - 2)
			End If 
		
			If InStr(part,"client_id") > 0 Then
				header = Split(part,"=")
				strResourceID = header(1)
				strResourceID = Mid(strResourceID,2,Len(strResourceID) - 2)
			End If 		
		Next
	
	End Function
	
	Private Function GetXDigestValue()
		
		Dim colElements
		oHTTP.open "POST", strSiteURL & "_api/contextinfo", False 
		oHTTP.setRequestHeader "accept","application/atom+xml;odata=verbose"
		oHTTP.setRequestHeader "authorization", "Bearer " & strSecurityToken
		oHTTP.send
		oXML.loadXML oHTTP.responseText
		Set colElements = oXML.getElementsByTagName("d:FormDigestValue")
		strFormDigestValue = colElements.item(0).text 
		
	End Function
	
	Private Function GetSecurityToken
	
		Dim oHTTP,part,parts,tokens,token
		Set oHTTP = CreateObject("MSXML2.ServerXMLHTTP")
		strURLbody = "grant_type=client_credentials&client_id=" & strClientID & "@" & strTenantID & "&client_secret=" & strClientSecret & "&resource=" & strResourceID & "/" & strDomain & "@" & strTenantID
		oHTTP.open "POST", strAuthUrlPart1 & strTenantID & strAuthUrlPart2, False
		oHTTP.setRequestHeader "User-Agent","Mozilla/5.0 (compatible; MSIE 9.0; Windows NT 6.1; WOW64; Trident/5.0)"
		oHTTP.setRequestHeader "Host","accounts.accesscontrol.windows.net"
		oHTTP.setRequestHeader "Content-Type","application/x-www-form-urlencoded"
		oHTTP.setRequestHeader "Content-Length", CStr(Len(strURLbody))
		oHTTP.send strURLbody
		'debug
'		debug.WriteLine "GetSecurityToken"
'		debug.WriteLine strURLbody
'		debug.WriteLine oHTTP.responseText
'		debug.WriteLine "-----------------------"
		'debug
		parts = Split(oHTTP.responseText,",")
		For Each part In parts
			If InStr(part,"access_token") > 0 Then
				tokens = Split(part,":")
				Exit For
			End If
		Next
		
		
		token = Mid(tokens(1),2,Len(tokens(1)) - 3)
		strSecurityToken = token
		
		
	End Function 
	
	Private Function Strip(sString)
		If Right(sString,1) = "/" Then
			sString = Mid(sString,1,Len(sString) - 1)
		End If 
		If Left(sString,1) = "/" Then
			sString = Mid(sString,2,Len(sString) - 1)
		End If
		
		Strip = sString
	End Function 
	
	Public Function GetListItem(sListName,sFieldName,sFieldValue)
		Dim oHTTP
		Set oHTTP = CreateObject("MSXML2.ServerXMLHTTP.3.0")
		With oHTTP
			.open "GET", strSiteURL & "_api/web/lists/GetByTitle('" & sListName & "')/items?$select=" & sFieldName & "&$filter=" & sFieldName & " eq " & "'" & sFieldValue & "'", False 
			.setRequestHeader "Authorization", "Bearer " & strSecurityToken
			.setRequestHeader "Accept", "application/atom+xml;odata=verbose"
			.send
		End With
		
		If Not oHTTP.status = 200 Then
			GetListItem = False ' Something went wrong. Assume the item doesn't exist an owervrite it. Or lose it !
			Exit Function
		End If 
		
		oXML.loadXML oHTTP.responseText
		
		If oXML.getElementsByTagName("d:Title").length > 0 Then
			If sFieldValue = oXML.getElementsByTagName("d:Title").item(0).text Then
				GetListItem = True
				Exit Function
			Else
				GetListItem = False
				Exit Function
			End If
		Else
			GetListItem = False
			Exit Function
		End If 
	End Function
	
	Public Function UpdateList()
	
		Dim oHTTP,body,oSTREAM
		body = "<?xml version=""1.0"" encoding=""utf-8""?>" _
		& "<soap12:Envelope xmlns:xsi=""http://www.w3.org/2001/XMLSchema-instance"" xmlns:xsd=""http://www.w3.org/2001/XMLSchema"" xmlns:soap12=""http://www.w3.org/2003/05/soap-envelope"">" _
		& "<soap12:Body><UpdateListItems xmlns=""http://schemas.microsoft.com/sharepoint/soap/"">" _
		& "<listName>SK01_Manual_Payments_QA</listName><updates><Field Name=""ID"">21<Field><Field Name=""Title"">HELLO</Field></updates></UpdateListItems></soap12:Body></soap12:Envelope>"
		Set oHTTP = CreateObject("MSXML2.ServerXMLHTTP.6.0")
		oHTTP.open "POST","https://volvogroup.sharepoint.com/sites/unit-rc-sk-bs-it/_vti_bin/Lists.asmx",False
		oHTTP.setRequestHeader "Host","volvogroup.sharepoint.com"
		oHTTP.setRequestHeader "Content-Type","application/soap+xml; charset=utf-8"
		oHTTP.setRequestHeader "Content-Length",Len(body)
		oHTTP.send body 
	End Function 
	
	Public Function GetFileInfo(sServerRelFilePath)
		
		If Not Left(sServerRelFilePath,1) = "/" Then
			sServerRelFilePath = "/" & sServerRelFilePath
		End If 
		 
		With oHTTP
			.open "GET", strSiteURL & "_api/web/getFileByServerRelativeUrl('/sites/" & strSite & sServerRelFilePath & "')/Properties"
			.setRequestHeader "Authorization", "Bearer " & strSecurityToken
			.setRequestHeader "Accept", "application/atom+xml;odata=verbose"
			.send
		End With 
		
		If Not oHTTP.status = 200 Then
			GetFileInfo = -1
			Exit Function
		End If 
		
		oXML.setProperty "SelectionNamespaces","xmlns=""http://www.w3.org/2005/Atom"""
		oXML.loadXML oHTTP.responseText
		debug.WriteLine oXML.getElementsByTagName("d:vti_x005f_filesize").item(0).text
		
	End Function 
	
			
			
	Public Function DownloadFile(sServerRelFilePath,sSaveAsPath)
	
		Dim oHTTP,oSTREAM
		Set oHTTP = CreateObject("MSXML2.ServerXMLHTTP")
		With oHTTP
			.open "GET", strSiteURL & "_layouts/15/download.aspx?SourceUrl=https://" & strDomain & sServerRelFilePath
			.setRequestHeader "Authorization", "Bearer " & strSecurityToken
			.send
		End With
		
		debug.WriteLine oHTTP.responseText
		debug.WriteLine oHTTP.status
		
		If oHTTP.status = 200 Then
			Set oSTREAM = CreateObject("ADODB.Stream")
			oSTREAM.Open
			oSTREAM.Type = 1
			oSTREAM.Write oHTTP.responseBody
			oSTREAM.SaveToFile sSaveAsPath
			oSTREAM.Close
		Else
			debug.WriteLine oHTTP.status
			debug.WriteLine oHTTP.responseText
		End If 
		
	End Function 
	
			
	Public Function GetFileCount(sServerRelDirPath)
	
		Dim oHTTP,oXML,colElements
		sServerRelDirPath = Strip(sServerRelDirPath)
		Set oHTTP = CreateObject("MSXML2.XMLHTTP")
		Set oXML = CreateObject("MSXML2.DOMDocument")
		
		With oHTTP
			.open "GET", oSP.GetSiteURL & "_api/web/GetFolderByServerRelativeUrl('/sites/" & strSite & sServerRelDirPath & "')/Files", False
			.setRequestHeader "Authorization", "Bearer " & oSP.GetToken
			.setRequestHeader "Accept", "application/atom+xml;odata=verbose"
			.send
		End With
		
		If Not oHTTP.status = 200 Then 
			GetFileCount = -1
			Exit Function
		End If 

		
		oXML.setProperty "SelectionNamespaces","xmlns=""http://www.w3.org/2005/Atom"""
		oXML.loadXML oHTTP.responseText
		Set colElements = oXML.getElementsByTagName("d:Name")
		
		GetFileCount = colElements.length
				
	End Function 
	
	Public Function FolderExists(sRelDirPath)
		
		Dim oHTTP,oXML,colElements
		Set oHTTP = CreateObject("MSXML2.XMLHTTP")
		Set oXML = CreateObject("MSXML2.DOMDocument")
		oXML.setProperty "SelectionNamespaces","xmlns=""http://www.w3.org/2005/Atom"""
		
		With oHTTP
			.open "GET", strSiteURL & "_api/web/GetFolderByServerRelativeUrl('" & sRelDirPath & "')", False
			.setRequestHeader "Authorization", "Bearer " & strSecurityToken
			.setRequestHeader "Accept", "application/atom+xml;odata=verbose"
			.send
		End With
		
		If Not oHTTP.status = 200 Then
			errSource = "MSXML2.XMLHTTP"
		End If 
		
		If oHTTP.status = 200 Then
			oXML.loadXML oHTTP.responseText
			Set colElements = oXML.getElementsByTagName("d:Exists")
			If colElements.length > 0 Then
				If LCase(colElements.item(0).text) = "true" Then
					FolderExists = True
					Exit Function
				Else
					FolderExists = False
					Exit Function
				End If 
			End If  
		Else
			FolderExists = False
		End If 
	
	End Function 
	
	Public Function DownloadFilesA(sServerRelDirPath,sDestinationFolder,ByRef arrFiles)
		
		Dim item,nodes
		Dim filesDownloaded : filesDownloaded = 0
		Dim path
		Dim oSTREAM : Set oSTREAM = CreateObject("ADODB.Stream")
		
		For item = 0 To UBound(arrFiles)
			
			With oHTTP
				.open "GET", strSiteURL & "_api/web/getFileByServerRelativeUrl('/sites/" & strSite & "/" & sServerRelDirPath & "/" & arrFiles(item) & "')/Properties", False
				.setRequestHeader "Authorization", "Bearer " & strSecurityToken
				.send
			End With 
			
			If Not oHTTP.status = 200 Then
				errSource = "MSXML2.XMLHTTP"
				errDescription = oHTTP.responseText
				errNumber = oHTTP.status
				DownloadFilesA = -1
				Exit Function 
			End If 
			
			With oHTTP
				.open "GET", strSiteURL & "_api/web/GetFolderByServerRelativeUrl('" & sServerRelDirPath & "')/Files('" & arrFiles(item) & "')", False
				.setRequestHeader "Authorization", "Bearer " & strSecurityToken
				.send
			End With
			
			
			If Not oHTTP.status = 200 Then
				errSource = "MSXML2.XMLHTTP"
				errDescription = oHTTP.responseText & "Affected file: " & arrFiles(item)
				errNumber = oHTTP.status
				DownloadFilesA = -1
				Exit Function
			End If 
			
			oXML.setProperty "SelectionNamespaces","xmlns=""http://www.w3.org/2005/Atom"""
			oXML.loadXML oHTTP.responseText
			
			Set nodes = oXML.getElementsByTagName("d:ServerRelativeUrl")
			
			If Not nodes.length > 0 Then
				errSource = "MSXML2.XMLHTTP"
				errDescription = "d:serverRelativeUrl node missing. Affected file: " & arrFiles(item)
				errNumber = Hex(1000)
				DownloadFilesA = -1
				Exit Function
			End If 
			
			path = nodes.nextNode.text ' Save relative URl
		
			With oHTTP
				.open "GET", strSiteURL & "_layouts/15/download.aspx?SourceUrl=https://" & strDomain & path, False
				.setRequestHeader "Authorization", "Bearer " & strSecurityToken
				.send
			End With
			
			If Not oHTTP.status = 200 Then
				errSource = "MSXML2.XMLHTTP"
				errDescription = oHTTP.responseText
				errNumber = oHTTP.status
				DownloadFilesA = -1
				Exit Function
			End If 
			
			
			oSTREAM.Open
			oSTREAM.Type = 1
			oSTREAM.Write oHTTP.responseBody
			On Error Resume Next 
			oSTREAM.SaveToFile oFSO.BuildPath(sDestinationFolder,oFSO.GetFileName(path)),2
			
			If err.number > 0 Then 
				errSource = err.Source
				errDescription = err.Description & " Affected file: " & oFSO.BuildPath(sDestinationFolder,oFSO.GetFileName(path))
				errNumber = err.number
				DownloadFilesA = -1
				oSTREAM.Close
				Exit Function
			End If 
			
			oSTREAM.Close
			filesDownloaded = filesDownloaded + 1
			
		Next
		
		DownloadFilesA = filesDownloaded
		
	End Function 
	
	
	
	
	
	
	Public Function DownloadFiles(sServerRelDirPath,sDestinationFolder)
	
		Dim item,nodes
		Dim path
		Dim oSTREAM : Set oSTREAM = CreateObject("ADODB.Stream")
		
		With oHTTP
			.open "GET", strSiteURL & "_api/web/GetFolderByServerRelativeUrl('/sites/" & strSite & sServerRelDirPath & "')/Files", False
			.setRequestHeader "Accept", "application/atom+xml;odata=verbose"
			.setRequestHeader "Authorization", "Bearer " & strSecurityToken
			.send
		End With 
		
		If Not oHTTP.status = 200 Then
			errDescription = oHTTP.responseText
			errNumber = oHTTP.status
			DownloadFiles = -1
			Exit Function
		End If 
		
		oXML.setProperty "SelectionNamespaces","xmlns=""http://www.w3.org/2005/Atom"""
		oXML.loadXML oHTTP.responseText
		
		Set nodes = oXML.getElementsByTagName("d:ServerRelativeUrl")
		
		For item = 0 To nodes.length - 1
			path = nodes.nextNode.text
			With oHTTP
				.open "GET", strSiteURL & "_layouts/15/download.aspx?SourceUrl=https://" & strDomain & path, False
				.setRequestHeader "Authorization", "Bearer " & strSecurityToken
				.send
			End With
			
			If Not oHTTP.status = 200 Then
				errDescription = oHTTP.responseText
				errNumber = oHTTP.status
				DownloadFiles = -1
				Exit Function
			End If 
			
			
			oSTREAM.Open
			oSTREAM.Type = 1
			oSTREAM.Write oHTTP.responseBody
			On Error Resume Next 
			oSTREAM.SaveToFile oFSO.BuildPath(sDestinationFolder,oFSO.GetFileName(path))
			
			If err.number > 0 Then 
				errSource = err.Source
				errDescription = err.Description
				errNumber = err.number
				DownloadFiles = -1
				Exit Function
			End If 
			
			oSTREAM.Close
			
		Next 
		
		DownloadFiles = nodes.length
		
	End Function 
	
	Public Function GetFilesA(sServerRelDirPath,ByRef dictFiles)
	
		Dim item,nodes
		With oHTTP
			.open "GET", strSiteURL & "_api/web/GetFolderByServerRelativeUrl('/sites/" & strSite & sServerRelDirPath & "')/Files", False
			.setRequestHeader "Authorization", "Bearer " & strSecurityToken
			.setRequestHeader "Accept", "application/atom+xml;odata=verbose"
			.send
		End With 
		
		If Not oHTTP.status = 200 Then
			GetFilesA = -1
			Exit Function
		End If 
		
		oXML.setProperty "SelectionNamespaces","xmlns=""http://www.w3.org/2005/Atom"""
		oXML.loadXML oHTTP.responseText
		
		Set nodes = oXML.getElementsByTagName("d:Name")
		
		For item = 0 To nodes.length - 1
			dictFiles.Add nodes.nextNode.text,""
		Next 
		
		GetFilesA = dictFiles.Count
	
	End Function 
		
	Public Function GetFiles(sServerRelDirPath,ByRef colFiles) ' sType "json" or "atom+xml"
	
		Dim dictFilesInSourceDir
		Dim dictFiles
		Dim oHTTP,oXML,colItems,item,colPaths,path
		Set dictFiles = CreateObject("Scripting.Dictionary")
		Set oXML = CreateObject("MSXML2.DOMDocument")
		Set oHTTP = CreateObject("MSXML2.ServerXMLHTTP.3.0")
		If Left(sServerRelDirPath,1) = "/" Then
			sServerRelDirPath = Right(sServerRelDirPath,Len(sServerRelDirPath) - 1)
		End If
		If Right(sServerRelDirPath,1) = "/" Then
			sServerRelDirPath = Left(sServerRelDirPath,Len(sServerRelDirPath) - 1)
		End If 
		
		With oHTTP
			.open "GET", strSiteURL & "_api/web/GetFolderByServerRelativeUrl('" & sServerRelDirPath & "')/Files", False
			.setRequestHeader "Authorization", "Bearer " & strSecurityToken
			.setRequestHeader "Accept", "application/atom+xml;odata=verbose"
			.send
		End With
	
		If oHTTP.status = 200 Then
			oXML.setProperty "SelectionNamespaces","xmlns=""http://www.w3.org/2005/Atom"""
			oXML.loadXML oHTTP.responseText
			Set colItems = oXML.getElementsByTagName("d:Name")
			Set colPaths = oXML.getElementsByTagName("d:ServerRelativeUrl")
		
			For item = 0 To colItems.length - 1
				colFiles.add colItems.item(item).text,colPaths.item(item).text
			Next
			
'			colFiles = dictFiles
'			Exit Function
		End If 
	
	End Function
			
			
	Public Function MoveFile2(sSourceRelDirPath,sDestRelDirPath)

		If Left(sSourceRelDirPath,1) = "/" Then
			sSourceRelDirPath = Right(sSourceRelDirPath,Len(sSourceRelDirPath) - 1)
		End If 
		If Left(sDestRelDirPath,1) = "/" Then
			sDestRelDirPath = "/" & Right(sDestRelDirPath,Len(sDestRelDirPath) - 1)
		End If 
		 
		Dim oHTTP,strBody
		Set oHTTP = CreateObject("MSXML2.ServerXMLHTTP.3.0")
		
		strBody = "{""srcPath"": {""__metadata"": {""type"": ""SP.ResourcePath""},""DecodedUrl"": """ & strSiteURL & sSourceRelDirPath & """},""destPath"": {""__metadata"": {""type"": ""SP.ResourcePath""},""DecodedUrl"": """ & strSiteURL & sDestRelDirPath & """}}"
		
		With oHTTP
			.open "POST","https://volvogroup.sharepoint.com/sites/unit-rc-sk-bs-it/_api/SP.MoveCopyUtil.MoveFileByPath(overwrite=@a1)?@a1=true"
			.setRequestHeader "Authorization", "Bearer " & strSecurityToken
			.setRequestHeader "Accept", "application/json;odata=nometadata"
			.setRequestHeader "Content-Type", "application/json;odata=verbose"
			.setRequestHeader "Content-Length", Len(strBody)
			.send strBody
		End With
		
		If oHTTP.status = 200 Then
			MoveFile2 = 1 ' Return 1 or True if successfull
			Exit Function
		Else
			MoveFile2 = 0 ' Return 0 or False if failed
			Exit Function
		End If 
	
	End Function
	
	
	Public Function AddListItem(sListName,sJsonRequest)
'		To do this operation, you must know the ListItemEntityTypeFullName property of the list And
'		pass that as the value of type in the HTTP request body. Following is a sample rest call to get the ListItemEntityTypeFullName

		Dim oHTTP,oXML,strEntityTypeFullName,colElements,request
		Set oHTTP = CreateObject("MSXML2.ServerXMLHTTP.3.0")
		Set oXML = CreateObject("MSXML2.DOMDocument")
		oXML.setProperty "SelectionNamespaces","xmlns=""http://www.w3.org/2005/Atom"""
		
		With oHTTP
			.open "GET", strSiteURL & "_api/web/lists/GetByTitle('" & sListName & "')?$select=ListItemEntityTypeFullName", False
			.setRequestHeader "Authorization", "Bearer " & strSecurityToken
			.setRequestHeader "Accept", "application/atom+xml;odata=verbose"
			.send
		End With 
		
		If oHTTP.status = 200 Then
			oXML.loadXML oHTTP.responseText
			Set colElements = oXML.getElementsByTagName("d:ListItemEntityTypeFullName")
			If colElements.length >= 1 Then
				strEntityTypeFullName = colElements.item(0).text
			Else
				AddListItem = -1 ' Couldn't obtain the EntityTypeFullName
				Exit Function 
			End If
		Else 
			AddListItem = -2 ' http error
			Exit Function 
		End If
		
		
		sJsonRequest = "{""__metadata"": { ""type"": """ & strEntityTypeFullName & """ }," & sJsonRequest ' Prepend metadata part
		With oHTTP
			.open "POST", strSiteURL & "_api/web/lists/GetByTitle('" & sListName & "')/items", False
			.setRequestHeader "Authorization", "Bearer " & strSecurityToken
			.setRequestHeader "Accept", "application/json;odata=verbose"
			.setRequestHeader "Content-Type", "application/json;odata=verbose"
			.setRequestHeader "If-None-Match", "*"
			.setRequestHeader "Content-Length", Len(sJsonRequest)
			.setRequestHeader "X-RequestDigest", strFormDigestValue
			.send sJsonRequest
		End With
		
		If oHTTP.status = 201 Then
			AddListItem = 0 ' Success
			Exit Function
		Else 
			AddListItem = oHTTP.status
			debug.WriteLine oHTTP.responseText
			Exit Function
		End If 
		
		
	End Function 
	
	Public Function DeleteAllItemsInList(sListName)
		
		Dim oHTTP,oXML,element
		Set oXML = CreateObject("MSXML2.DOMDocument")
		Set oHTTP = CreateObject("MSXML2.ServerXMLHTTP.3.0")
		
		With oHTTP
			.open "GET", strSiteURL & "_api/web/lists/GetByTitle('" & sListName & "')/items", False 
			.setRequestHeader "Authorization", "Bearer " & strSecurityToken
			.setRequestHeader "Accpet", "application/json;odata=verbose"
			.setRequestHeader "Content-Type", "application/json"
'			.setRequestHeader "If-Match", "{etag or *}"
'			.setRequestHeader "X-HTTP-Method", "DELETE"
			.send
		End With 
		
		oXML.loadXML oHTTP.responseText
		For Each element In oXML.getElementsByTagName("d:Id")
			With oHTTP
				.open "POST", strSiteURL & "_api/web/lists/GetByTitle('" & sListName & "')/items(" & element.text & ")", False
				.setRequestHeader "Authorization", "Bearer " & strSecurityToken
				.setRequestHeader "Accpet", "application/json;odata=verbose"
				.setRequestHeader "Content-Type", "application/json"
				.setRequestHeader "If-Match", "*"
				.setRequestHeader "X-HTTP-Method", "DELETE"
				.send
			End With 
		Next 
	End Function 
		
			
	
	
	
	
	' ************************** P R O P E R T I E S ******************************
	Public Property Get LastErrorSource
	
		LastErrorSource = errSource
		
	End Property 
	
	Public Property Get LastErrorCode
	
		LastErrorCode = errNumber
		
	End Property
	
	Public Property Get LastErrorDesc
	
		LastErrorDesc = errDescription
		
	End Property 
	
	Public Property Get GetDigest
		
		GetDigest = strFormDigestValue
		
	End Property 
	
	Public Property Get GetToken
	
		GetToken = strSecurityToken
		
	End Property 
	
	Public Property Get GetHttpResponse
		GetHttpResponse = oHTTP.responseText
	End Property 
	
	Public Property Get GetHttpResponseHeaders(strHeader) ' If strHeader "*" then get all headers
		If strHeader = "*" Then
		
			GetHttpResponseHeaders = oHTTP.getAllResponseHeaders
			Exit Property
		
		End If
		
		GetHttpResponseHeaders = oHTTP.getResponseHeader(strHeader)
		
	End Property 
	
	Public Property Get GetRealmTenantID
		GetRealmTenantID = strTenantID
	End Property
	
	Public Property Get GetClientID
		GetClientID = strClientID
	End Property 
	
	Public Property Get GetResourceID
		GetResourceID = strResourceID
	End Property 
	
	Public Property Get GetClientSecret
		GetClientSecret = strClientSecret
	End Property 
	
	Public Property Get GetAuthURL
		GetAuthURL = strAuthUrlPart1 & strTenantID & strAuthUrlPart2
	End Property 
	
	Public Property Get GetSiteURL
		GetSiteURL = strSiteURL
	End Property 
	
	Public Property Get GetSiteDomain
		GetSiteDomain = strDomain
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
		
		Set oSAPGUI = GetObject("SAPGUI") ' This fails is saplogon is not running. We're connecting to the COM object not creating our own instance
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