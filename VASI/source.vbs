Option Explicit
'#######################################
'############ Constants ################
'####################################### 
' Excel worksheet 'OUT' column names
Const DOC_TYPE = "A"  
Const BREAK_DOWN_PROD_LINE = "B"
Const DOCUMENT_OUT_NUMBER = "C"
Const REQUEST_NUMBER = "D"
Const CUSTOMER_NAME = "E"
Const PROCESSED_YEAR = "F"
Const PROCESSED_WEEK = "G"
Const VIN_SERIAL = "H"
Const REGISTRATION = "I"
Const DEALER_ID = "J"
Const DEALER_PARMA_NO = "K"
Const DEALER_NAME = "L"
Const DOC_DATE = "M"
Const DEFAULT_TEMPLATE = "N"
Const CURRENCY_CODE = "O"
Const TOTAL_AMOUNT = "P"
Const LABOUR = "Q"
Const TRAVEL_TIME = "R"
Const BRANDED_PARTS = "S"
Const DEALER_COUNTRY = "T"
Const NON_BRANDED_PARTS = "U"
Const ENV_SUNDRY = "V"
Const DEALER_CALLOUT = "W"
Const CCP = "X"
Const TOWING_COST = "Y"
Const SERVICE_VAN = "Z"
Const OTHER_COST = "AA"
Const THIRD_PARTY_WORK = "AB"
Const SUBTOTAL = "AC"
Const CHARGED_FOREIGN_VAT = "AD"
Const TOTAL_ADMIN_CHARGE = "AE"
Const TOTAL_BELGIAN_VAT = "AF"
Const VAT_CHARGEABLE = "AG"
Const NON_VAT_CHARGEABLE = "AH"
Const BREAKDOWN_DATE = "AI"
' Excel worksheet 'CCP' column names
Const DOC_TYPE_CCP = "A"
Const REQUEST_NUMBER_CCP = "B"
Const BREAK_DOWN_PROD_LINE_CCP = "C"
Const DOCUMENT_OUT_NUMBER_CCP = "E"
Const BREAKDOWN_DATE_CCP = "F"
Const TOTAL_AMOUNT_CCP = "G"
Const CURRENCY_CODE_CCP = "H"
Const DEALER_ID_CCP = "I"
Const DEALER_NAME_CCP = "J"
Const DEALER_COUNTRY_CCP = "K"
' Sharepoint list column names <-----------------------> Excel sheet column names
Const S_BREAKDOWN_DATE = "BreakdownDate" ' <-----------> BREAKDOWN_DATE
Const S_CURRENCY_CODE = "CurrencyCode" ' <-----------------> CURRENCY_CODE
Const S_PDF = "PDF"
Const S_INVOICE = "Invoice" ' <------------------------> DOCUMENT_OUT_NUMBER
Const S_REQUEST_NUMBER = "RequestNumber" ' <-----------> REQUEST_NUMBER
Const S_CUSTOMER_NAME = "CustomerName" ' <-------------> CUSTOMER_NAME
Const S_YEAR_WEEK = "ProcessedYear" ' <----------------> PROCESSED_YEAR + PROCESSED_WEEK
Const S_VIN_SERIAL = "VINSerial" ' <-------------------> VIN_SERIAL
Const S_REG_NO = "RegNo" ' <---------------------------> REGISTRATION
Const S_DEALER_ID = "DealerId" ' <---------------------> Lookup DEALER_ID in the dealers dictionary
Const S_DOC_DATE = "DocDate" ' <-----------------------> DOC_DATE
Const S_CONTRACT = "Contract" ' <----------------------> Lookup contract in the contracts file
Const S_DEFAULT_TEMPLATE = "DefTempl" ' <--------------> DEFAULT_TEMPLATE
Const S_CURRENCY = "Currency" ' <----------------------> CURRENCY_CODE
Const S_TOTAL_AMOUNT = "TotalAmount" ' <---------------> TOTAL_AMOUNT
Const S_LABOUR = "Labour" ' <--------------------------> LABOUR
Const S_TRAVEL_TIME = "TravelTime" ' <-----------------> TRAVEL_TIME
Const S_BRANDED_PARTS = "BrandedParts" ' <-------------> BRANDED_PARTS
Const S_DEALER_COUNTRY = "DealerCountry" ' <-----------> DEALER_COUNTRY
Const S_NON_BRANDED_PARTS = "NonBrandedParts" ' <------> NON_BRANDED_PARTS
Const S_ENV_SUNDRY = "EnvSundry" ' <-------------------> ENV_SUNDRY
Const S_DEALER_CALLOUT = "DealerCallOut" ' <-----------> DEALER_CALLOUT
Const S_CCP = "CCP" ' <--------------------------------> CCP
Const S_TOWING_COST = "TowingCost" ' <-----------------> TOWING_COST
Const S_SERVICE_VAN = "ServiceVan" ' <-----------------> SERVICE_VAN
Const S_OTHER_COST = "OtherCost" ' <-------------------> OTHER_COST
Const S_THIRD_PARTY_WORK = "ThirdPartyWork" ' <--------> THIRD_PARTY_WORK
Const S_SUBTOTAL = "SubTotal" ' <----------------------> SUBTOTAL
Const S_CHARGED_FOREIGN_VAT = "ChargedForeignVAT" ' <--> CHARGED_FOREIGN_VAT
Const S_TOTAL_ADMIN_CHARGE = "TotalAdminCharge" ' <----> TOTAL_ADMIN_CHARGE
Const S_TOTAL_BELGIAN_VAT = "TotalBelgianVAT" ' <------> TOTAL_BELGIAN_VAT
Const S_VAT_CHARGEABLE = "VATChargeable" ' <-----------> VAT_CHARGEABLE
Const S_NON_VAT_CHARGEABLE = "NonVATChargeable" ' <----> NON_VAT_CHARGEABLE
Const S_BRAND = "Brand" ' <----------------------------> BREAK_DOWN_PROD_LINE
Const S_REGISTRATION = "RegNo" ' <---------------------> REGISTRATION
Const S_DOCTYPE = "DocType" ' <------------------------> DOCTYPE

' Other constants
Const outdir = "CZ02_VASI" ' Output directory
Const spsourcedir = "CZ02_VASI/SOURCE" ' Sharepoint online source directory
Const spprocesseddir = "CZ02_VASI/SOURCE/PROCESSED" ' Sharepoint online processed directory
Const attachmentdir = "CZ02_VASI/CZ02_VASI_PDF" ' Sharepoint online attachments directory
Const CONTRACTS = "CZ02_VASI_Contracts.xlsx"
Const SP_LIST_NAME = "CZ02_VASI_Portal"
'###########################################
'############## Variables ##################
'###########################################
Dim s__pdf, s__invoice, s__requestnumber, s__customername, s__yearweek, s__vinserial, s__regno, s__dealerid, s__docdate
Dim s__contract, s__deftempl, s__currency, s__totalamount, s__labour, s__traveltime, s__brandedparts, s__nonbrandedparts
Dim s__envsundry, s__dealercallout, s__ccp, s__towingcost, s__servicevan, s__othercost, s__thirdpartywork, s__subtotal
Dim s__processedyear, s__chargedforeignvat, s__totaladmincharge, s__totalbelgianvat, s__vatchargeable, s__nonvatchargeable
Dim s__brand, s__documentoutnumber, s__dealercountry, s__breakdowndate, s__doctype
Dim cstart,cexp,dd,ctype,temp,bd
Dim strSiteFullUrl
Dim outpath,strDate,strBreakDate,strRequest
Dim dictFilesInSourceDir,dictFilesInAttachmentDir,dictDealerLoc
Dim numRows,row,files,file,drive,vinnum
Dim oSP,oHTTP,oFSO,oWSH,oFILE,oSTREAM,oXML,oCON,oRST,oEXCEL,oWB,oWS,oRX,oWBC,oWSC,oRNG,oWSCCCP
Set oRX = New RegExp
Set oWSH = CreateObject("Wscript.Shell")
Set dictDealerLoc = CreateObject("Scripting.Dictionary")
Set dictFilesInSourceDir = CreateObject("Scripting.Dictionary")
Set dictFilesInAttachmentDir = CreateObject("Scripting.Dictionary")
Set oFSO = CreateObject("Scripting.FileSystemObject")
Set oEXCEL = Wscript.CreateObject("Excel.Application","ExcelOn_")
Set oXML = CreateObject("MSXML2.DOMDocument")
Set oSTREAM = CreateObject("ADODB.Stream")
Set oRST = CreateObject("ADODB.RecordSet")
Set oCON = CreateObject("ADODB.Connection")
Set oHTTP = CreateObject("MSXML2.XMLHTTP")
Set oSP = New SP
oRX.IgnoreCase = True
oRX.Global = True
dictDealerLoc.Add "24270","VTC"
dictDealerLoc.Add "27728","NYR"
dictDealerLoc.Add "28129","BRN"
dictDealerLoc.Add "28130","HRA"
dictDealerLoc.Add "28131","UST"
dictDealerLoc.Add "28132","CBU"
dictDealerLoc.Add "28133","HAV"
dictDealerLoc.Add "28134","OTR"
dictDealerLoc.Add "28135","LOD"
dictDealerLoc.Add "28136","HUM"
dictDealerLoc.Add "28137","OLO"
dictDealerLoc.Add "28138","TUR"
dictDealerLoc.Add "28140","POP"
dictDealerLoc.Add "28142","STR"
dictDealerLoc.Add "28143","KLE"
dictDealerLoc.Add "024270","VTC"
dictDealerLoc.Add "027728","NYR"
dictDealerLoc.Add "028129","BRN"
dictDealerLoc.Add "028130","HRA"
dictDealerLoc.Add "028131","UST"
dictDealerLoc.Add "028132","CBU"
dictDealerLoc.Add "028133","HAV"
dictDealerLoc.Add "028134","OTR"
dictDealerLoc.Add "028135","LOD"
dictDealerLoc.Add "028136","HUM"
dictDealerLoc.Add "028137","OLO"
dictDealerLoc.Add "028138","TUR"
dictDealerLoc.Add "028140","POP"
dictDealerLoc.Add "028142","STR"
dictDealerLoc.Add "028143","KLE"
dictDealerLoc.Add "311625","NYR"
dictDealerLoc.Add "311502","HRA"
dictDealerLoc.Add "311594","UST"
dictDealerLoc.Add "312765","CBU"
dictDealerLoc.Add "311992","OTR"
dictDealerLoc.Add "312691","LOD"
dictDealerLoc.Add "310735","HUM"
dictDealerLoc.Add "311414","OLO"
dictDealerLoc.Add "312218","TUR"
dictDealerLoc.Add "306520","RTC"
dictDealerLoc.Add "313417","OST"
dictDealerLoc.Add "313492","POP"
dictDealerLoc.Add "320059","STR"
dictDealerLoc.Add "321708","KLE"
'dictDealerLoc.Add "25249","SC"
'dictDealerLoc.Add "311508","SC"
'dictDealerLoc.Add "27747","ZA"
'dictDealerLoc.Add "312046","ZA"
'dictDealerLoc.Add "27748","TO"
'dictDealerLoc.Add "312051","TO"
'dictDealerLoc.Add "27749","PO"
'dictDealerLoc.Add "312044","PO"
'dictDealerLoc.Add "27750","ZV"
'dictDealerLoc.Add "311529","ZV"
'dictDealerLoc.Add "27751","NZ"
'dictDealerLoc.Add "312045","NZ"
'dictDealerLoc.Add "27752","PP"
'dictDealerLoc.Add "318688","PP"

'#########################################################################################################
'################################################# M A I N ###############################################
'#########################################################################################################

'*******************************************************************
' Pre-processing tasks
' Creating folder structure
'*******************************************************************
oSP.Init "https://volvogroup.sharepoint.com/sites/unit-vasi","volvogroup.sharepoint.com","40784cc3-ba68-45d0-9891-f3dfa8f04d15","cDES2gLLi%2BBRI/FcUizAyZuGQFQ5p%2B6rrknc3kMBWmE="
' Check if Sharepoint folder exists
If Not oSP.FolderExists(spsourcedir) Then
	'ERROR. Folder doesn't exist
	'SEND MESSAGE
	WScript.Quit
End If 

outpath = Left(oWSH.SpecialFolders("desktop"),1) & ":\!AUTO" ' Get system drive (C)

If Not oFSO.FolderExists(outpath) Then
	oFSO.CreateFolder outpath
	If err.number <> 0 Then
		'ERROR. Folder can't be created
		'SEND MESSAGE
		WScript.Quit
	End If 
End If 

If Not oFSO.FolderExists(outpath & "\" & outdir) Then
	oFSO.CreateFolder outpath & "\" & outdir
	If err.number <> 0 Then
		'ERROR. Folder can't be created
		'SEND MESSAGE
		WScript.Quit
	End If 
	outpath = outpath & "\" & outdir & "\"
Else
	outpath = outpath & "\" & outdir & "\"
End If 

strSiteFullUrl = "https://" & oSP.GetSiteDomain
'*********************************************************************
' Download attachments files (PDF)
' Download source file (excel)
' Download contracts file (excel)
'*********************************************************************
debug.WriteLine "Searching " & oSP.GetSiteURL & attachmentdir
oSP.GetFiles attachmentdir,dictFilesInAttachmentDir ' Function returns dictionary of all files in the specified directory
debug.WriteLine "Searching " & oSP.GetSiteURL & spsourcedir
oSP.GetFiles spsourcedir,dictFilesInSourceDir ' Function returns dictionary of all the files in the specified directory so that we can process each file
oRX.Pattern = "[0-9]{4}[W|M][0-9]{2}"
For Each file In dictFilesInSourceDir.Keys
	If oFSO.FileExists(outpath & file) Then
		oFSO.DeleteFile outpath & file
	End If 
	debug.WriteLine "Downloading file " & file
	oSP.DownloadFile dictFilesInSourceDir.Item(file), outpath & file
Next
If Not oFSO.FileExists(outpath & CONTRACTS) Then
	debug.WriteLine "Contracts file not found"
	'SEND MESSAGE
	WScript.Quit
End If 
'*******************************************************************
' Process downloaded source files (excel)
' Open Contracts file (excel)
'*******************************************************************
Set oWBC = oEXCEL.Workbooks.Open(outpath & CONTRACTS)
Set oWSC = oWBC.Worksheets("Report 1")
For Each file In dictFilesInSourceDir.Keys
' dorobit nacitavanie subor _M_
	oRX.Pattern = "[0-9]{4}[W|M][0-9]{2}"
	If oRX.Test(file) Then
		debug.WriteLine "Processing file " & file
		Set oWB = oEXCEL.Workbooks.Open(outpath & file) ' Open local file
		Set oWS = oWB.Worksheets("OUT")
		Set oWSCCCP = oWB.Worksheets("CCP")
		numRows = oWS.Usedrange.Rows.count
		row = 2
		debug.WriteLine "Processing workbook " & oWB.Name & " sheet " & oWS.Name
		debug.WriteLine "Rows in the " & oWS.Name & " " & numRows
		Do While row <= numRows
			If Not oSP.GetListItem(SP_LIST_NAME,"Title",Trim(oWS.Range(DOCUMENT_OUT_NUMBER & row).Value)) Then 
				If dictFilesInAttachmentDir.Exists(oWS.Range(DOCUMENT_OUT_NUMBER & row).Value & ".pdf") Then ' We have the PDF file
					s__pdf = """" & S_PDF & """:{""__metadata"":{""type"":""SP.FieldUrlValue""},""Description"":""@"",""Url"":""https://" & oSP.GetSiteDomain & dictFilesInAttachmentDir.Item(oWS.Range(DOCUMENT_OUT_NUMBER & row).Value & ".pdf") & """ }," ' Link to the attached PDF
				Else
					s__pdf = """" & S_PDF & """:{""__metadata"":{""type"":""SP.FieldUrlValue""},""Description"":""-"",""Url"":null }," ' Link to the attached PDF
				End If
				s__documentoutnumber = oWS.Range(DOCUMENT_OUT_NUMBER & row).Value
				s__requestnumber = oWS.Range(REQUEST_NUMBER & row).Value
				strDate = Month(CDate(oWS.Range(DOC_DATE & row).Value)) & "." & Day(CDate(oWS.Range(DOC_DATE & row).Value)) & "." & Year(CDate(oWS.Range(DOC_DATE & row).Value))
				strBreakDate = oWS.Range(BREAKDOWN_DATE & row).Value
				'************************************************************************
				' Determine if the vehicle has an active contract based on BREAKDOWN_DATE
				' Set ctype appropriately
				'************************************************************************
				vinnum = oWS.Range(VIN_SERIAL & row).Value ' Save the vin number
				
				On Error Resume Next
					Set oRNG = oWSC.Range("A2:A" & oWSC.UsedRange.Rows.Count).Find(vinnum) ' Should not generate error
					cstart = oWSC.Range("B" & oRNG.Row).Value ' Accessing variable that is not an object should generate error
				If err.number <> 0 Then ' This VIN has no value in the contract file
					debug.WriteLine "VIN " & vinnum & " not found in the contracts file"
					ctype = "-" ' No contract
				Else ' This VIN has a value in the contract file
					debug.WriteLine "VIN " & vinnum & " found in the contracts file"
					cstart = CLng(Year(CDate(oWSC.Range("B" & oRNG.Row).Value)) & Right("00" & Month(CDate(oWSC.Range("B" & oRNG.Row).Value)),2) & Right("00" & Day(CDate(oWSC.Range("B" & oRNG.Row).Value)),2))
					cexp = CLng(Year(CDate(oWSC.Range("C" & oRNG.Row).Value)) & Right("00" & Month(CDate(oWSC.Range("B" & oRNG.Row).Value)),2) & Right("00" & Day(CDate(oWSC.Range("B" & oRNG.Row).Value)),2))
					bd = CLng(Year(CDate(oWS.Range(BREAKDOWN_DATE & row).Value)) & Right("00" & Month(CDate(oWS.Range(BREAKDOWN_DATE & row).Value)),2) & Right("00" & Day(CDate(oWS.Range(BREAKDOWN_DATE & row).Value)),2))
					If bd >= cstart And bd <= cexp Then
						ctype = oWSC.Range("E" & oRNG.Row).Value
						debug.WriteLine vbTab & ctype & " contract was valid at the time"
					Else
						debug.WriteLine vbTab & ctype & " contract was invalid at the time"
						ctype = "-" ' Contract was not valid at the time
					End If 
				End If 
				On Error GoTo 0
				
				If LCase(oWS.Range(DEFAULT_TEMPLATE & row).Value) = "true" Then
					s__deftempl = """" & S_DEFAULT_TEMPLATE & """:""Yes"","
				Else
					s__deftempl = """" & S_DEFAULT_TEMPLATE & """:""No"","
				End If
				s__doctype = """" & S_DOCTYPE & """:""" & oWS.Range(DOC_TYPE & row).Value & ""","
				s__regno = """" & S_REGISTRATION & """:""" & oWS.Range(REGISTRATION & row).Value & ""","
				s__dealercallout = """" & S_DEALER_CALLOUT & """:""" & Replace(oWS.Range(DEALER_CALLOUT & row).Value,",",".") & ""","
				s__contract = """" & S_CONTRACT & """:""" & ctype & ""","
				s__breakdowndate = """" & S_BREAKDOWN_DATE & """:""" & Replace(strBreakDate,"/",".") & ""","
				s__docdate = """" & S_DOC_DATE & """:""" & strDate & ""","
				s__documentoutnumber = """Title"":""" & oWS.Range(DOCUMENT_OUT_NUMBER & row).Value & ""","
				s__requestnumber = """" & S_REQUEST_NUMBER & """:""" & oWS.Range(REQUEST_NUMBER & row).Value & ""","
				s__customername = """" & S_CUSTOMER_NAME & """:""" & oWS.Range(CUSTOMER_NAME & row).Value & ""","
				s__yearweek = """" & S_YEAR_WEEK & """:""" & oWS.Range(PROCESSED_YEAR & row).Value & "-" & Right("00" & oWS.Range(PROCESSED_WEEK & row).Value,2) & ""","
				s__vinserial = """" & S_VIN_SERIAL & """:""" & oWS.Range(VIN_SERIAL & row).Value & ""","
				s__dealerid = """" & S_DEALER_ID & """:""" & dictDealerLoc.Item(oWS.Range(DEALER_ID & row).Value) & "-" & oWS.Range(DEALER_ID & row).Value & ""","
				s__currency = """" & S_CURRENCY_CODE & """:""" & oWS.Range(CURRENCY_CODE & row).Value & ""","
				s__totalamount = """" & S_TOTAL_AMOUNT & """:""" & Replace(Replace(oWS.Range(TOTAL_AMOUNT & row).Value,",","."),"-","0") & ""","
				s__labour = """" & S_LABOUR & """:""" & Replace(Replace(oWS.Range(LABOUR & row).Value,",","."),"-","0") & ""","
				s__traveltime = """" & S_TRAVEL_TIME & """:""" & Replace(Replace(oWS.Range(TRAVEL_TIME & row).Value,",","."),"-","0") & ""","
				s__brandedparts = """" & S_BRANDED_PARTS & """:""" & Replace(Replace(oWS.Range(BRANDED_PARTS & row).Value,",","."),"-","0") & ""","
				s__dealercountry = """" & S_DEALER_COUNTRY & """:""" & Replace(oWS.Range(DEALER_COUNTRY & row).Value,",","0") & ""","
				s__nonbrandedparts = """" & S_NON_BRANDED_PARTS & """:""" & Replace(Replace(oWS.Range(NON_BRANDED_PARTS & row).Value,",","."),"-","0") & ""","
				s__envsundry = """" & S_ENV_SUNDRY & """:""" & Replace(Replace(oWS.Range(ENV_SUNDRY & row).Value,",","."),"-","0") & ""","
				s__dealercallout = """" & S_DEALER_CALLOUT & """:""" & Replace(Replace(oWS.Range(DEALER_CALLOUT & row).Value,",","."),"-","0") & ""","
				s__ccp = """" & S_CCP & """:""" & Replace(Replace(oWS.Range(CCP & row).Value,",","."),"-","0") & ""","
				s__towingcost = """" & S_TOWING_COST & """:""" & Replace(Replace(oWS.Range(TOWING_COST & row).Value,",","."),"-","0") & ""","
				s__servicevan = """" & S_SERVICE_VAN & """:""" & Replace(Replace(oWS.Range(SERVICE_VAN & row).Value,",","."),"-","0") & ""","
				s__othercost = """" & S_OTHER_COST & """:""" & Replace(Replace(oWS.Range(OTHER_COST & row).Value,",","."),"-","0") & ""","
				s__thirdpartywork = """" & S_THIRD_PARTY_WORK & """:""" & Replace(Replace(oWS.Range(THIRD_PARTY_WORK & row).Value,",","."),"-","0") & ""","
				s__subtotal = """" & S_SUBTOTAL & """:""" & Replace(Replace(oWS.Range(SUBTOTAL & row).Value,",","."),"-","0") & ""","
				s__chargedforeignvat = """" & S_CHARGED_FOREIGN_VAT & """:""" & Replace(Replace(oWS.Range(CHARGED_FOREIGN_VAT & row).Value,",","."),"-","0") & ""","
				s__totaladmincharge = """" & S_TOTAL_ADMIN_CHARGE & """:""" & Replace(Replace(oWS.Range(TOTAL_ADMIN_CHARGE & row).Value,",","."),"-","0") & ""","
				s__totalbelgianvat = """" & S_TOTAL_BELGIAN_VAT & """:""" & Replace(Replace(oWS.Range(TOTAL_BELGIAN_VAT & row).Value,",","."),"-","0") & ""","
				s__vatchargeable = """" & S_VAT_CHARGEABLE & """:""" & Replace(Replace(oWS.Range(VAT_CHARGEABLE & row).Value,",","."),"-","0") & ""","
				If LCase(oWS.Range(BREAK_DOWN_PROD_LINE & row).Value) = "volvo" Then
					s__brand = """" & S_BRAND & """:""VT"","
				ElseIf LCase(oWS.Range(BREAK_DOWN_PROD_LINE & row).Value) = "renault" Then
					s__brand = """" & S_BRAND & """:""RT"","
				End If 
				s__nonvatchargeable = """" & S_NON_VAT_CHARGEABLE & """:""" & Replace(oWS.Range(NON_VAT_CHARGEABLE & row).Value,",",".") & """}"
				strRequest = s__pdf & s__documentoutnumber & s__requestnumber & s__contract & s__customername & s__currency & s__yearweek & s__vinserial & s__dealerid _
						   & s__docdate & s__totalamount & s__labour & s__regno & s__traveltime & s__brandedparts & s__dealercountry & s__deftempl & s__doctype _
						   & s__nonbrandedparts & s__envsundry & s__dealercallout & s__ccp & s__towingcost & s__servicevan & s__othercost & s__thirdpartywork _
						   & s__subtotal & s__chargedforeignvat & s__totaladmincharge & s__totalbelgianvat & s__breakdowndate & s__vatchargeable & s__brand & s__nonvatchargeable
						   
				debug.WriteLine oSP.AddListItem("CZ02_VASI_Portal",strRequest)
			Else
				Debug.WriteLine Trim(oWS.Range(DOCUMENT_OUT_NUMBER & row).Value) & " already exists in the sharepoint list " & SP_LIST_NAME & ". Skipping this item"
			End If 
			row = row + 1
		Loop
		'*******************
		'Process CCP sheet
		'*******************
		debug.WriteLine "Processing workbook " & oWB.Name & " sheet " & oWSCCCP.Name
		numRows = oWSCCCP.Usedrange.Rows.count
		debug.WriteLine "Rows in the " & oWSCCCP.Name & " " & numRows
		row = 2
		Do While row <= numRows
			If Not oSP.GetListItem(SP_LIST_NAME,"Title",Trim(oWSCCCP.Range(DOCUMENT_OUT_NUMBER_CCP & row).Value)) Then 
				If dictFilesInAttachmentDir.Exists(oWSCCCP.Range(DOCUMENT_OUT_NUMBER_CCP & row).Value & ".pdf") Then ' We have the PDF file
					s__pdf = """" & S_PDF & """:{""__metadata"":{""type"":""SP.FieldUrlValue""},""Description"":""@"",""Url"":""https://" & oSP.GetSiteDomain & dictFilesInAttachmentDir.Item(oWSCCCP.Range(DOCUMENT_OUT_NUMBER_CCP & row).Value & ".pdf") & """ }," ' Link to the attached PDF
				Else
					s__pdf = """" & S_PDF & """:{""__metadata"":{""type"":""SP.FieldUrlValue""},""Description"":""-"",""Url"":null }," ' Link to the attached PDF
				End If
				
				If LCase(oWS.Range(BREAK_DOWN_PROD_LINE_CCP & row).Value) = "volvo" Then
					s__brand = """" & S_BRAND & """:""VT"","
				ElseIf LCase(oWSCCCP.Range(BREAK_DOWN_PROD_LINE_CCP & row).Value) = "renault" Then
					s__brand = """" & S_BRAND & """:""RT"","
				End If 
				
				s__dealercountry = """" & S_DEALER_COUNTRY & """:""" & Replace(oWSCCCP.Range(DEALER_COUNTRY_CCP & row).Value,",","0") & ""","
				s__dealerid = """" & S_DEALER_ID & """:""" & dictDealerLoc.Item(oWSCCCP.Range(DEALER_ID_CCP & row).Value) & "-" & oWSCCCP.Range(DEALER_ID_CCP & row).Value & ""","
				s__currency = """" & S_CURRENCY_CODE & """:""" & oWSCCCP.Range(CURRENCY_CODE_CCP & row).Value & ""","
				s__doctype = """" & S_DOCTYPE & """:""" & oWSCCCP.Range(DOC_TYPE_CCP & row).Value & ""","
				s__requestnumber = """" & S_REQUEST_NUMBER & """:""" & oWSCCCP.Range(REQUEST_NUMBER_CCP & row).Value & ""","
				s__breakdowndate = """" & S_BREAKDOWN_DATE & """:""" & Replace(strBreakDate,"/",".") & ""","
				s__documentoutnumber = """Title"":""" & oWSCCCP.Range(DOCUMENT_OUT_NUMBER_CCP & row).Value & ""","
				s__totalamount = """" & S_TOTAL_AMOUNT & """:""" & Replace(Replace(oWSCCCP.Range(TOTAL_AMOUNT_CCP & row).Value,",","."),"-","0") & """}"
				
				strRequest = s__pdf & s__brand & s__dealercountry & s__dealerid & s__currency & s__doctype & s__requestnumber & s__breakdowndate & s__documentoutnumber & s__totalamount
				
				debug.WriteLine oSP.AddListItem("CZ02_VASI_Portal",strRequest)
			Else
				Debug.WriteLine Trim(oWSCCCP.Range(DOCUMENT_OUT_NUMBER_CCP & row).Value) & " already exists in the sharepoint list " & SP_LIST_NAME & ". Skipping this item"
			End If 
			
			row = row + 1
		Loop
		debug.WriteLine "Closing file " & file
		oWB.Close
	End If 
Next
debug.WriteLine "Closing file " & CONTRACTS
oWBC.Close
'**************** Move from SOURCE to PROCESSED ******************************
For Each file In dictFilesInSourceDir.Keys
	If Not file = CONTRACTS Then 
		debug.WriteLine "Moving file " & file & " from " & spsourcedir & " to " & spprocesseddir
		oSP.MoveFile2 spsourcedir & "/" & file, spprocesseddir & "/" & file
	End If 
Next 

For Each file In dictFilesInSourceDir.Keys
	If oRX.Test(file) Then
		If oFSO.FileExists(outpath & file) Then
			debug.WriteLine "Deleting file " & outpath & file
			oFSO.DeleteFile outpath & file
		End If 
	End If 
Next

If oFSO.FileExists(outpath & contracts) Then
	debug.WriteLine "Deleting file " & outpath & contracts
	oFSO.DeleteFile outpath & contracts
End If 


WScript.Quit
'#########################################################################################################
'########################################## M A I N  E N D  ##############################################
'#########################################################################################################



'#######################################
'## C L A S S   D E F I N I T I O N S ##
'#######################################

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
	
	Private Sub Class_Initialize
		
		numHTTPstatus = 0
		strAuthUrlPart1 = "https://accounts.accesscontrol.windows.net/"
		strAuthUrlPart2 = "/tokens/OAuth/2"
		Set oHTTP = CreateObject("MSXML2.ServerXMLHTTP.3.0")
		Set oXML = CreateObject("MSXML2.DOMDocument")
		
	End Sub 
	
	Public Function Init(sSiteUrl,sDomain,sClientID,sClientSecret)
	
		If Right(sSiteUrl,1) = "/" Then
			strSiteURL = sSiteUrl
		Else
			strSiteURL = sSiteUrl & "/"
		End If 
		
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
	
		
	
	'******************** P U B L I C   F U N C T I O N S ***********************
'	Public Function AddListItem(sListName,sJsonInput)
'	
'	
'		
'	End Function 
'	
'	Public Function CreateFolder(strParentFolder,strFolder)
'		
'		oHTTP.open "POST", strSiteURL & "_api/web/GetFolderByServerRelativeUrl('" & Replace(strParentFolder," ","%20") & "')/Folders/add('" & Replace(strFolder," ","%20") & "')", False
'		With oHTTP
'			.setRequestHeader "Accept", "application/json;odata=verbose"
'			.setRequestHeader "Content-Type", "application/json;odata=verbose"
'			.setRequestHeader "Authorization", "Bearer " & strSecurityToken
'			.setRequestHeader "X-RequestDigest", strFormDigestValue
'			.send
'		End With
'		
'		If oHTTP.status <> 200 Then
'			CreateFolder = oHTTP.status
'			Exit Function
'		End If 
'			
'		CreateFolder = 200

'	End Function 
'	
'	Public Function DeleteFolder(strParentFolder,strFolder)
'		
'		oHTTP.open "POST", strSiteURL & "_api/web/GetFolderByServerRelativeUrl('" & Replace(strParentFolder," ","%20") & "/" & strFolder & "')", False
'		With oHTTP
'			.setRequestHeader "Accept", "application/json;odata=verbose"
'			.setRequestHeader "Content-Type", "application/json;odata=verbose"
'			.setRequestHeader "Authorization", "Bearer " & strSecurityToken
'			.setRequestHeader "If-Match","*"
'			.setRequestHeader "X-HTTP-Method","DELETE"
'			.setRequestHeader "X-RequestDigest", strFormDigestValue
'			.send
'		End With
'		
'		If oHTTP.status <> 200 Then
'			DeleteFolder = oHTTP.status
'			Exit Function
'		End If 
'			
'		DeleteFolder = 200
'		
'	End Function 
'	
'	Public Function SPAddListItem(strListName,strFieldName,strFieldValue)
'	
'	End Function 
'	
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
	
	
	Public Function DownloadFile(sServerRelFilePath,sSaveAsPath)
	
		Dim oHTTP
		Set oHTTP = CreateObject("MSXML2.ServerXMLHTTP")
		With oHTTP
			.open "GET", strSiteURL & "_layouts/15/download.aspx?SourceUrl=https://" & strDomain & sServerRelFilePath
			.setRequestHeader "Authorization", "Bearer " & strSecurityToken
			.send
		End With
		
		If oHTTP.status = 200 Then
			Set oSTREAM = CreateObject("ADODB.Stream")
			oSTREAM.Open
			oSTREAM.Type = 1
			oSTREAM.Write oHTTP.responseBody
			oSTREAM.SaveToFile sSaveAsPath
			oSTREAM.Close
		End If
		
	End Function 
	
	Public Function GetFileCountXML(sServerRelDirPath,boolRetColObject,ByRef colObject)
	
		Dim oHTTP,oXML,colElements
		sServerRelDirPath = Strip(sServerRelDirPath)
		Set oHTTP = CreateObject("MSXML2.XMLHTTP")
		Set oXML = CreateObject("MSXML2.DOMDocument")
		
		With oHTTP
			.open "GET", oSP.GetSiteURL & "_api/web/GetFolderByServerRelativeUrl('" & sServerRelDirPath & "')/Files", False
			.setRequestHeader "Authorization", "Bearer " & oSP.GetToken
			.setRequestHeader "Accept", "application/atom+xml;odata=verbose"
			.send
		End With

		If oHTTP.status = 200 Then
			oXML.setProperty "SelectionNamespaces","xmlns=""http://www.w3.org/2005/Atom"""
			oXML.loadXML oHTTP.responseText
			Set colElements = oXML.getElementsByTagName("d:Name")
			If Not boolRetColObject Then ' Return file count only
				GetFileCount = colElements.length ' return the file count
				Exit Function
			Else 
				Set colObject = colElements
				GetFileCount = colElements.length
				Exit Function
			End If 
		End If 
		
		GetFileCount = -1 ' Return error
			
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
			
	
'	Public Function MoveFile(sSourceRelDirPath,sDestRelDirPath)
'	
'		Dim oHTTP
'		Set oHTTP = CreateObject("MSXML2.ServerXMLHTTP.3.0")
'		
'		With oHTTP
'			.open "GET", strSiteURL & "_api/web/GetFileByServerRelativeUrl('" & sSourceRelDirPath & "')/moveto(newurl='" & sDestRelDirPath & "',flags=1)"
'			.setRequestHeader "Authorization", "Bearer " & strSecurityToken
'			.send
'		End With 
'		
'		If oHTTP.status = 200 Then
'			MoveFile = 0
'			Exit Function
'		Else
'			debug.WriteLine oHTTP.status
'			debug.WriteLine oHTTP.responseText
'			MoveFile = -1
'			Exit Function
'		End If 
'	
'	End Function
	
	Public Function MoveFile2(sSourceRelDirPath,sDestRelDirPath)

		If Left(sSourceRelDirPath,1) = "/" Then
			sSourceRelDirPath = Right(sSourceRelDirPath,Len(sSourceRelDirPath) - 1)
		End If 
		If Left(sDestRelDirPath,1) = "/" Then
			sDestRelDirPath = "/" & Right(sDestRelDirPath,Len(sDestRelDirPath) - 1)
		End If 
		 
		Dim oHTTP,strBody
		Set oHTTP = CreateObject("MSXML2.ServerXMLHTTP.3.0")
		'strBody = "{ ""srcPath"": { ""__metadata"": ""SP.ResourcePath"" },""DecodeUrl"":" & strSiteURL & sSourceRelDirPath & """},""destPath"": { ""__metadata"": ""SP.ResourcePath"" },""DecodeUrl"":" & strSiteURL & sDestRelDirPath & """ } }"
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
			debug.WriteLine "HTTP->OK"
			Exit Function
		Else 
			AddListItem = oHTTP.status
			debug.WriteLine "HTTP-> " & oHTTP.responseText
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
	Public Property Get Errcode
	
		Errcode = numHTTPstatus
		
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



Function ExcelOn_Open(bCancel)
	debug.WriteLine "BeforeClose event fired"
End Function 
