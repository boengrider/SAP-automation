#AutoIt3Wrapper_UseX64=Y
#include <StringConstants.au3>
#include <WindowsConstants.au3>
#include <AutoItConstants.au3>
#include <String.au3>
#include <WinAPISys.au3>
#include <APIConstants.au3>
#include <FileConstants.au3>
#include <WinAPIFiles.au3>



#Region About
;Example Call #1
;         -s fq2        -c 105          -src C:\!AUTO\SI01_HADES_DND_NET      -cc si01            -nal																 -nosub																		-as https://XXXXXXXXXXX.sharepoint.com/sites/XXXXXXXXXX/SI01_HADES_ARCHIVE
;         ^^^^^^^^^^    ^^^^^^^^^^^     ^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^      ^^^^^^^^^^^^^       ^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^    ^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^       ^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^
;         SAP SYSTEM    SAP CLIENT      Source location				           Company code       No local archivation (if ommited files will be archived localy      Do not process subfolders (if ommited subfolders will be processed)        Arhive files to the sharepoint site
;Example Call #2
;		  -s fq2 -c 105 -src https://XXXXXXXXX.sharepoint.com/sites/XXXXXXX/SI01_HADES_SOURCE -cc si01 -nal -nosub -as https://XXXXXXXXXXXX/sites/XXXXXXXXXXXX/SI01_HADES_ARCHIVE
;Exit codes
;1-99 - Something other than errors e.g no parameters passed
;100 - General script errors
;200 - SAP related errors
;300 - Sharepoint related errors
#EndRegion

#Region Main script

#Region Constants
Enum $SAP_APP = 0, $SAP_CON, $SAP_SES
Enum $CRED_USER = 0, $CRED_PASSWORD, $CRED_DOMAIN, $CRED_HOST; 0,1,2,3
Enum $CLI_SYSTEM = 0, $CLI_CLIENT, $CLI_ARCHIVESHAREPOINT, $CLI_ARCHIVESHAREPOINTURL, $CLI_SOURCE, $CLI_ARCHIVELOCAL, $CLI_COMPANYCODE, $CLI_SUBDIRS ; 0,1,2,3,4,5,6,7
Enum $SOURCE_UNKNOWN = 0, $SOURCE_NETDRIVE, $SOURCE_SHAREPOINT; 0,1,2 Source type
Enum $FILE_PARENTFOLDER = 0, $FILE_PATH, $FILE_NAME, $FILE_SIZE, $FILE_SOURCETYPE, $FILE_STATUS; File status, File parent subfolder in case name is constructed from subfolder and file names, Full file path or URL if it's sharepoint , Source type i.e net drive or sharepoint
Enum $STATE_INTIAL = 0, $STATE_COPYBEGIN, $STATE_ALREADYEXISTS, $STATE_COPIED, $STATE_DRAGDROPREADY, $STATE_DRAGDROPPED, $STATE_OK, $STATE_INVALID, $STATE_PROCESSED, $STATE_DRAGDROPBEGIN, $STATE_DRAGDROPFAILED, $STATE_ARCHIVELOCALBEGIN, $STATE_ARCHIVEDLOCAL, $STATE_ARCHIVESPBEGIN, $STATE_ARCHIVEDSP, $STATE_DELETEREMOTEBEGIN, $STATE_DELETEDREMOTE
;Const $MAX_SIZEINBYTES = -1 ; Unlimited
Const $MAX_SIZEINBYTES = 10485760
Const $MANDATORY_PARAMS = -4
Const $RESOURCE_NAME = "SPRESTAPI"
Const $PROJECT = "_HADES_DND" ; Company code is prepended
Const $AUTO_FOLDER_ROOT = "C:\!AUTO"
Const $AUTO_FOLDER_SOURCE = "SOURCE"
Const $AUTO_FOLDER_ARCHIVE = "ARCHIVE"
Const $SAP_LOCAL_LANDSCAPE_PATH = @AppDataDir & "\SAP\Common\SAPUILandscape.xml"
Const $SYS_ADMINS = "jon.doe@company.com;jane.doe@company.com"
#EndRegion

#Region Variables
Local $oMyError = ObjEvent("AutoIt.Error", "ErrFunc")
Local $RunTimeStamp = @MDAY & @MON & @YEAR & "_" & @HOUR & @MIN & @SEC
Local $FileCount = -1
Local $DataSource = 0
Local $ExitCode = 0
Local $SapGuiApplication = Null
Local $SapGuiConnection = Null
Local $PidSapLogon
Local $SapSystemDescription = Null
Local $SpAccessToken = Null
Local $LogFile = Null
Local $SysInfo = ObjCreate("ADSystemInfo")
Local $Fso = ObjCreate("Scripting.FileSystemObject")
Local $User = ObjGet("LDAP://" & $SysInfo.UserName)
Local $Net = ObjCreate("Wscript.Network")
Local $Http = ObjCreate("winhttp.winhttprequest.5.1")
Local $FilesToProcess = ObjCreate("Scripting.Dictionary")
Local $OawdFolders = ObjCreate("Scripting.Dictionary")
$OawdFolders.Add("CZ02", "CZ02 FI Scanning")
$OawdFolders.Add("SI01", "SI01 FI Scan")
Local $SAP[3] = [Null,Null,Null]
Local $Credentials[4] = [Null,Null,Null,Null] ; Array holding credentials
Local $CliParams[8] = [Null,Null,False,Null,Null,True,Null,True] ; Array holding command line parameters. For order of values refer to the $CLI Enum
Local $FileState[17] = ["INTIAL", "COPYBEGIN", "ALREADYEXISTS", "COPIED", "DRAGDROPREADY", "DRAGDROPPED", "VALID", "INVALID", "PROCESSED", "DRAGDROPBEGIN", "DRAGDROPFAILED", "ARCHIVELOCALBEGIN", "ARCHIVEDLOCAL", "ARCHIVESPBEGIN", "ARCHIVEDSP", "DELETEREMOTEBEGIN", "DELETEDREMOTE"]
Local $ActualParamsPassed = 1 ; Initially 1. Shift to the right each time a mandatory parameter is present. We should end up with 1 >> NUMBER_OF_MANDATORY_PARAMETERS
#EndRegion


#Region IsLocked
If IsWorkstationLocked() = 0 Then
   MessageToAdmin("W;" & @ScriptName & ";" & @YEAR & "-" & @MON & "-" & @MDAY & ";" & @HOUR & ":" & @MIN & ":" & @SEC & ";" & @UserName & ";" & @ComputerName & ";" & $CliParams[$CLI_SYSTEM],"Workstation is locked. Connect to the HADES workstation and login/unlock user profile.",$SYS_ADMINS, Null)
   Exit(99)
EndIf
#EndRegion





#Region Process command line parameters
For $arg = 1 To $CmdLine[0]

   Switch StringUpper($CmdLine[$arg])

   Case "-NOSUB"
	  $CliParams[$CLI_SUBDIRS] = False

   Case "--NOSUBDIR"
	  $CliParams[$CLI_SUBDIRS] = False

   Case "-CC"
	  If $CmdLine[0] > $arg Then
		 $CliParams[$CLI_COMPANYCODE] = $CmdLine[$arg + 1]
		 $ActualParamsPassed = BitShift($ActualParamsPassed,-1)
	  EndIf

   Case "-COMPANYCODE"
	  If $CmdLine[0] > $arg Then
		 $CliParams[$CLI_COMPANYCODE] = $CmdLine[$arg + 1]
		 $ActualParamsPassed = BitShift($ActualParamsPassed,-1)
	  EndIf

   Case "-S"
	  If $CmdLine[0] > $arg Then
		 $CliParams[$CLI_SYSTEM] = $CmdLine[$arg + 1]
		 $ActualParamsPassed = BitShift($ActualParamsPassed,-1)
	  EndIf

   Case "--SYSTEM"
	  If $CmdLine[0] > $arg Then
		 $CliParams[$CLI_SYSTEM] = $CmdLine[$arg + 1]
		 $ActualParamsPassed = BitShift($ActualParamsPassed,-1)
	  EndIf

   Case "-C"
	  If $CmdLine[0] > $arg Then
		 $CliParams[$CLI_CLIENT] = $CmdLine[$arg + 1]
		 $ActualParamsPassed = BitShift($ActualParamsPassed,-1)
	  EndIf

   Case "--CLIENT"
	  If $CmdLine[0] > $arg Then
		 $CliParams[$CLI_CLIENT] = $CmdLine[$arg + 1]
		 $ActualParamsPassed = BitShift($ActualParamsPassed,-1)
	  EndIf

   Case "-AS"
	  $CliParams[$CLI_ARCHIVESHAREPOINT] = True
	  If $CmdLine[0] > $arg Then
		 If StringRegExp($CmdLine[$arg + 1], "(https?:\/\/(?:www\.|(?!www))[a-zA-Z0-9][a-zA-Z0-9-]+[a-zA-Z0-9]\.[^\s]{2,}|www\.[a-zA-Z0-9][a-zA-Z0-9-]+[a-zA-Z0-9]\.[^\s]{2,}|https?:\/\/(?:www\.|(?!www))[a-zA-Z0-9]+\.[^\s]{2,}|www\.[a-zA-Z0-9]+\.[^\s]{2,})",$STR_REGEXPMATCH) Then
			$CliParams[$CLI_ARCHIVESHAREPOINTURL] = $CmdLine[$arg + 1]
		 EndIf
	  EndIf

   Case "--ARCHIVESHAREPOINT"
	  $CliParams[$CLI_ARCHIVESHAREPOINT] = True
	  If $CmdLine[0] > $arg Then
		 If StringRegExp($CmdLine[$arg + 1], "(https?:\/\/(?:www\.|(?!www))[a-zA-Z0-9][a-zA-Z0-9-]+[a-zA-Z0-9]\.[^\s]{2,}|www\.[a-zA-Z0-9][a-zA-Z0-9-]+[a-zA-Z0-9]\.[^\s]{2,}|https?:\/\/(?:www\.|(?!www))[a-zA-Z0-9]+\.[^\s]{2,}|www\.[a-zA-Z0-9]+\.[^\s]{2,})",$STR_REGEXPMATCH) Then
			$CliParams[$CLI_ARCHIVESHAREPOINTURL] = $CmdLine[$arg + 1]
		 EndIf

	  EndIf

   Case "-SRC"
	  If $CmdLine[0] > $arg Then
		 If StringRegExp($CmdLine[$arg + 1], "(https?:\/\/(?:www\.|(?!www))[a-zA-Z0-9][a-zA-Z0-9-]+[a-zA-Z0-9]\.[^\s]{2,}|www\.[a-zA-Z0-9][a-zA-Z0-9-]+[a-zA-Z0-9]\.[^\s]{2,}|https?:\/\/(?:www\.|(?!www))[a-zA-Z0-9]+\.[^\s]{2,}|www\.[a-zA-Z0-9]+\.[^\s]{2,})",$STR_REGEXPMATCH) _
			Or StringRegExp($CmdLine[$arg + 1], "(\\{2}[a-zA-Z0-9]*){1}\\.*", $STR_REGEXPMATCH) Or StringRegExp($CmdLine[$arg + 1], "^[a-zA-Z]:\\[\\\S|*\S]?.*$", $STR_REGEXPMATCH) Then
		    $CliParams[$CLI_SOURCE] = $CmdLine[$arg + 1]
			$ActualParamsPassed = BitShift($ActualParamsPassed,-1)
		 EndIf
	  EndIf

   Case "--SOURCE"
	  If $CmdLine[0] > $arg Then
		 If StringRegExp($CmdLine[$arg + 1], "(https?:\/\/(?:www\.|(?!www))[a-zA-Z0-9][a-zA-Z0-9-]+[a-zA-Z0-9]\.[^\s]{2,}|www\.[a-zA-Z0-9][a-zA-Z0-9-]+[a-zA-Z0-9]\.[^\s]{2,}|https?:\/\/(?:www\.|(?!www))[a-zA-Z0-9]+\.[^\s]{2,}|www\.[a-zA-Z0-9]+\.[^\s]{2,})",$STR_REGEXPMATCH) _
			Or StringRegExp($CmdLine[$arg + 1], "(\\{2}[a-zA-Z0-9]*){1}\\.*", $STR_REGEXPMATCH) Or StringRegExp($CmdLine[$arg + 1], "^[a-zA-Z]:\\[\\\S|*\S]?.*$", $STR_REGEXPMATCH) Then
		    $CliParams[$CLI_SOURCE] = $CmdLine[$arg + 1]
			$ActualParamsPassed = BitShift($ActualParamsPassed,-1)
		 EndIf
	  EndIf

   Case "-NAL"
	  $CliParams[$CLI_ARCHIVELOCAL] = False

   Case "--NOARCHLOCAL"
	  $CliParams[$CLI_ARCHIVELOCAL] = False

   EndSwitch

Next

;There are 4 mandatory parameters at this point
If $ActualParamsPassed <> BitShift(1,$MANDATORY_PARAMS) Then
   ConsoleWrite("Usage: " & @ScriptName & " -s | --system $SAP_SYSTEM_NAME -c | --client $SAP_CLIENT_NAME -cc | --companycode $COMPANY_CODE -src | --source $SHAREPOINT_URL_OR_NETWORK_SHARE_PATH [-as | --archivesharepoint] [-nal | --noarchlocal] [-nosub | --nosubdir]" & @CRLF)
   Exit(1)
Else
   ConsoleWrite("--------------------------------------------------" & @CRLF)
   ConsoleWrite("Command line parameters summary" & @CRLF)
   ConsoleWrite("--------------------------------------------------" & @CRLF)
   ConsoleWrite("Company code: " & $CliParams[$CLI_COMPANYCODE] & @CRLF)
   ConsoleWrite("SAP System: " & $CliParams[$CLI_SYSTEM] & @CRLF)
   ConsoleWrite("SAP Client: " & $CliParams[$CLI_CLIENT] & @CRLF)
   ConsoleWrite("Source: " & $CliParams[$CLI_SOURCE] & @CRLF)
   ConsoleWrite("Process sub directories: " & $CliParams[$CLI_SUBDIRS] & @CRLF)
   ConsoleWrite("Archive to Sharepoint: " & $CliParams[$CLI_ARCHIVESHAREPOINT] & " ( " & $CliParams[$CLI_ARCHIVESHAREPOINTURL] & " )" & @CRLF)
   ConsoleWrite("Archive locally: " & $CliParams[$CLI_ARCHIVELOCAL] & @CRLF)
EndIf

#EndRegion


#Region Build local folder structure
;DirCreate(@AppDataDir & "\HADES\" & StringUpper($CliParams[$CLI_COMPANYCODE]))
;$LogFile = FileOpen(@AppDataDir & "\HADES\" & StringUpper($CliParams[$CLI_COMPANYCODE]) & "\" & $RunTimeStamp & ".log.txt", $FO_APPEND)
DirCreate($AUTO_FOLDER_ROOT & "\" & StringUpper($CliParams[$CLI_COMPANYCODE]) & $PROJECT)
DirCreate($AUTO_FOLDER_ROOT & "\" & StringUpper($CliParams[$CLI_COMPANYCODE]) & $PROJECT & "\" & StringUpper($CliParams[$CLI_COMPANYCODE]) & $PROJECT & "_" & $AUTO_FOLDER_SOURCE)
DirCreate($AUTO_FOLDER_ROOT & "\" & StringUpper($CliParams[$CLI_COMPANYCODE]) & $PROJECT & "\" & StringUpper($CliParams[$CLI_COMPANYCODE]) & $PROJECT & "_" & $AUTO_FOLDER_ARCHIVE)
$LogFile = FileOpen($AUTO_FOLDER_ROOT & "\" & StringUpper($CliParams[$CLI_COMPANYCODE]) & $PROJECT & "\" & $RunTimeStamp & ".log.txt", $FO_APPEND)
$LogFilePath = $AUTO_FOLDER_ROOT & "\" & StringUpper($CliParams[$CLI_COMPANYCODE]) & $PROJECT & "\" & $RunTimeStamp & ".log.txt"
#EndRegion

#Region Log CLI parameters
LogEvent($LogFile, "******************************************", False)
LogEvent($LogFile, "Script run: " & $RunTimeStamp, False)
LogEvent($LogFile, "******************************************", False)
LogEvent($LogFile, "CLI parameters", False)
LogEvent($LogFile, "Company code: " & $CliParams[$CLI_COMPANYCODE], False)
LogEvent($LogFile, "SAP System: " & $CliParams[$CLI_SYSTEM], False)
LogEvent($LogFile, "SAP Client: " & $CliParams[$CLI_CLIENT], False)
LogEvent($LogFile, "Source: " & $CliParams[$CLI_SOURCE], False)
LogEvent($LogFile, "Process sub directories: " & $CliParams[$CLI_SUBDIRS], False)
LogEvent($LogFile, "Archive to Sharepoint: " & $CliParams[$CLI_ARCHIVESHAREPOINT] & " ( " & $CliParams[$CLI_ARCHIVESHAREPOINTURL] & " )", False)
LogEvent($LogFile, "Archive locally: " & $CliParams[$CLI_ARCHIVELOCAL], False)
#EndRegion

#Region Get credentials, obtain sharepoint access token
; Get token here even if the source is not a sharepoint since we will need the token for Wdapp and possible sharepoint archive and Wdapp
$Credentials = GetCredentials($RESOURCE_NAME)

$SpAccessToken = SPGetAccessToken($Http, $Credentials[$CRED_DOMAIN], $Credentials[$CRED_HOST], $Credentials[$CRED_USER], $Credentials[$CRED_PASSWORD])

If @error Then
   $ExitCode = @error
   LogEvent($LogFile, "Error obtaining Sharepoint access token", True)
   ConsoleWrite("Error obtaining access token (" & @error & ")" & @CRLF)
   Exit($ExitCode)
Endif

ConsoleWrite("Sharepoint access token: " & $SpAccessToken & @CRLF)
#EndRegion


#Region check for residual file(s)
Local $ResidualFiles = DirGetSize($AUTO_FOLDER_ROOT & "\" & StringUpper($CliParams[$CLI_COMPANYCODE]) & $PROJECT & "\" & StringUpper($CliParams[$CLI_COMPANYCODE]) & $PROJECT & "_" & $AUTO_FOLDER_SOURCE, $DIR_EXTENDED)
If $ResidualFiles[1] > 0 Then
   LogEvent($LogFile, "Residual file (" & $ResidualFiles[1] & ") found in the local source location " & $AUTO_FOLDER_ROOT & "\" & StringUpper($CliParams[$CLI_COMPANYCODE]) & $PROJECT & "\" & StringUpper($CliParams[$CLI_COMPANYCODE]) & $PROJECT & "_" & $AUTO_FOLDER_SOURCE, True)
    MessageToAdmin("E;" & @ScriptName & ";" & $CliParams[$CLI_SYSTEM],"Residual file(s) found in the local source. See the attached log file", $SYS_ADMINS, $LogFilePath)
	Exit(98)
EndIf
#EndRegion


#Region Verify source location
; Source is a local fodler or a unc path (net share)
If StringRegExp($CliParams[$CLI_SOURCE], "(\\{2}[a-zA-Z0-9]*){1}\\.*") Or StringRegExp($CliParams[$CLI_SOURCE], "^[a-zA-Z]:\\[\\\S|*\S]?.*$", $STR_REGEXPMATCH) Then
   If Not FileExists($CliParams[$CLI_SOURCE]) Then
	  LogEvent($LogFile, "Hades source folder does not exist", True)
	  Exit(101) ; Source folder does not exist
   EndIf

   $FileCount = NSFolderGetFileCount($CliParams[$CLI_SOURCE], $CliParams[$CLI_SUBDIRS], $FilesToProcess)
   If @error Then
	  Exit(@error)
   EndIf
   $DataSource = $SOURCE_NETDRIVE
; Source is a sharepoint site
ElseIf StringRegExp($CliParams[$CLI_SOURCE], "(https?:\/\/(?:www\.|(?!www))[a-zA-Z0-9][a-zA-Z0-9-]+[a-zA-Z0-9]\.[^\s]{2,}|www\.[a-zA-Z0-9][a-zA-Z0-9-]+[a-zA-Z0-9]\.[^\s]{2,}|https?:\/\/(?:www\.|(?!www))[a-zA-Z0-9]+\.[^\s]{2,}|www\.[a-zA-Z0-9]+\.[^\s]{2,})", $STR_REGEXPMATCH) Then
   $FileCount = SPFolderGetFileCount($Http, $CliParams[$CLI_SOURCE], $SpAccessToken, $CliParams[$CLI_SUBDIRS])
   If @error Then
	  Exit(@error) ; Sharepoint library/folder does not exist
   ElseIf $FileCount < 0 Then
	  Exit(102) ; Unkown error while getting file count in the SP library
   EndIf
   $DataSource = $SOURCE_SHAREPOINT
EndIf

If $FileCount = 0 Then
   LogEvent($LogFile, "Nothing to process in the source location " & $CliParams[$CLI_SOURCE], True)
   MessageToAdmin("I;" & @ScriptName & ";" & $CliParams[$CLI_SYSTEM],"No files for processing found in the location: " & $CliParams[$CLI_SOURCE],"tomas.ac@volvo.com", $LogFilePath)
   Exit(2) ; Nothing to process
EndIf

#EndRegion


#Region SAP initialization

If Not SAPLaunchSAPLogon($PidSapLogon) Then
   ConsoleWrite("SAP Logon couldn't be launched")
   Exit(200) ; SAP Logon couldn't be launched
EndIf

ConsoleWrite("SAP Logon launched. PID -> " & $PidSapLogon & @CRLF)

$SapGuiConnection = SAPOpenConnection($CliParams[$CLI_SYSTEM], $CliParams[$CLI_CLIENT], $Net.UserName)

If $SapGuiConnection == Null Or Not IsObj($SapGuiConnection) Then
   ConsoleWrite("SAP Connection to the system " & SAPGetSystemDescription($CliParams[$CLI_SYSTEM]) & " couldn't be obtained" & @CRLF)
   Exit(202)
EndIf

Local $SessionCount = $SAP[$SAP_CON].Sessions.Count
Local $SessionOAWD = $SAP[$SAP_SES] ; Primary session created with first connection
Local $SessionCWID

LogEvent($LogFile, "SAP session count: " & $SessionCount, False)
LogEvent($LogFile, "Executing transaction '" & $SAP[$SAP_SES].Info.Transaction & "'", False)

If StringRegExp(StringUpper($SAP[$SAP_SES].Info.Transaction), "S000") Then
   $SAP[$SAP_CON].CloseConnection
   LogEvent($LogFile, "Failed to login: " & $Net.UserName & " into " & $CliParams[$CLI_SYSTEM] & " " & $CliParams[$CLI_CLIENT], True)
   MessageToAdmin("E;" & @ScriptName & ";" & $CliParams[$CLI_SYSTEM], "Failed to login: " & $Net.UserName & " into " & $CliParams[$CLI_SYSTEM] & " " & $CliParams[$CLI_CLIENT], $SYS_ADMINS, $LogFilePath)
EndIf

LogEvent($LogFile, "Creating a new session", False)
$SessionOAWD.CreateSession ; Secondary session
While $SAP[$SAP_CON].Sessions.Count = $SessionCount
   Sleep(100)
WEnd

$SessionCount = $SAP[$SAP_CON].Sessions.Count
$SessionCWID = $SAP[$SAP_CON].Sessions.Item($SessionCount - 1)
LogEvent($LogFile, "SAP session count: " & $SessionCount, False)

$SessionOAWD.StartTransaction("OAWD")
SAPKillPopups($SessionOAWD)
LogEvent($LogFile, "Executing transaction '" & $SessionOAWD.Info.Transaction & "' in the session '" & $SessionOAWD.Name, False)
$SessionCWID.StartTransaction("zfidocwid")
SAPKillPopups($SessionCWID)
LogEvent($LogFile, "Executing transaction '" & $SessionCWID.Info.Transaction & "' in the session '" & $SessionCWID.Name, False)

$SessionOAWD.findById("wnd[0]").sendVKey(71)	;Ctrl+F to find string
$SessionOAWD.findById("wnd[1]/usr/txtRSYSF-STRING").text = $OawdFolders.Item(StringUpper($CliParams[$CLI_COMPANYCODE]))
$SessionOAWD.findById("wnd[1]").sendVKey(0)	;Enter key
$SessionOAWD.findById("wnd[2]").sendVKey(84)	;Ctrl+G to point to the searched string
$SessionOAWD.findById("wnd[2]").sendVKey(2)	;F2 key to select the pointed string and return to previous wnd
$SessionOAWD.findById("wnd[0]").sendVKey(2)	;F2 key to expand the pointed position
$SessionOAWD.findById("wnd[0]").sendVKey(71)	;Ctrl+F to find string
$SessionOAWD.findById("wnd[1]/usr/txtRSYSF-STRING").text = "Incoming invoice prel posting (PDF)"	;string in find control
$SessionOAWD.findById("wnd[1]").sendVKey(0)	;Enter key
$SessionOAWD.findById("wnd[2]").sendVKey(84)	;Ctrl+G to point to the searched string
$SessionOAWD.findById("wnd[2]").sendVKey(2)	;F2 key to select the pointed string and return to previous wnd
$SessionOAWD.findById("wnd[0]").sendVKey(2)	;F2 key to expand the pointed position
$SessionOAWD.findById("wnd[1]/usr/txtCONFIRM_DATA-NOTE").text = "" ;nazov PDF suboru bez PDF a max 50 znakov

Local $WinCWID = $SessionCWID.ActiveWindow
Local $WinOAWD = $SessionOAWD.findById("wnd[1]")
#EndRegion


#Region explorer window open
Local $PathSrc = $AUTO_FOLDER_ROOT & "\" & StringUpper($CliParams[$CLI_COMPANYCODE]) & $PROJECT & "\" & StringUpper($CliParams[$CLI_COMPANYCODE]) & $PROJECT & "_" & $AUTO_FOLDER_SOURCE
Local $PathArch = $AUTO_FOLDER_ROOT & "\" & StringUpper($CliParams[$CLI_COMPANYCODE]) & $PROJECT & "\" & StringUpper($CliParams[$CLI_COMPANYCODE]) & $PROJECT & "_" & $AUTO_FOLDER_ARCHIVE
Local $ExplorerTitle = StringUpper($CliParams[$CLI_COMPANYCODE]) & $PROJECT & "_" & $AUTO_FOLDER_SOURCE

;First kill all the windows with source folder opened so that we work with the only one window and all key commands are recevied by this window
While WinExists($ExplorerTitle,"")
   WinKill($ExplorerTitle,"")
WEnd

If IsExplorerOpen($PathSrc, $ExplorerTitle) Then
   LogEvent($LogFile, "Explorer window is opened. Setting the view to Extra Large Icons", False)
   Sleep(3000)
   ExplorerSetView($ExplorerTitle)
Else
   LogEvent($LogFile, "Explorer window is opened. Setting the view to Extra Large Icons", False)
   If Not OpenWindowsExplorer($PathSrc) Then
	  LogEvent($LogFile, "Unable to open source directory " & $PathSrc & " in the windows explorer", True)
	  MessageToAdmin("E;" & @ScriptName & ";" & $CliParams[$CLI_SYSTEM],"", $SYS_ADMINS, $LogFilePath)
	  Exit(120)
   EndIf

   Sleep(3000)
   LogEvent($LogFile, "Explorer window is opened. Setting the view to Extra Large Icons", False)
   ExplorerSetView($ExplorerTitle)
EndIf

ExplorerSetView($ExplorerTitle)
#EndRegion

#Region State machine
Local $DropResult
Local $Timer = TimerInit()
Local $File
Local $ExplorerWindowTitle = StringUpper($CliParams[$CLI_COMPANYCODE]) & $PROJECT & "_" & $AUTO_FOLDER_SOURCE
Local $LocalArchiveFolder = $AUTO_FOLDER_ROOT & "\" & StringUpper($CliParams[$CLI_COMPANYCODE]) & $PROJECT & "\" & StringUpper($CliParams[$CLI_COMPANYCODE]) & $PROJECT & "_" & $AUTO_FOLDER_ARCHIVE & "\" & $RunTimeStamp
Local $LocalSourceFolder = $AUTO_FOLDER_ROOT & "\" & StringUpper($CliParams[$CLI_COMPANYCODE]) & $PROJECT & "\" & StringUpper($CliParams[$CLI_COMPANYCODE]) & $PROJECT & "_" & $AUTO_FOLDER_SOURCE
;$STATE_INTIAL = 0, $STATE_COPYBEGIN, $STATE_ALREADYEXISTS, $STATE_COPIED, $STATE_DRAGDROP, $STATE_DROPPED, $STATE_OK, $STATE_INVALID, $STATE_PROCESSED
LogEvent($LogFile, "-------------------------------------------------------", False)
For $File In $FilesToProcess.Items
   ; **********************************************************************
   ;Before each loop there should be no file in the local source directory
   ; **********************************************************************
   $ResidualFiles = DirGetSize($AUTO_FOLDER_ROOT & "\" & StringUpper($CliParams[$CLI_COMPANYCODE]) & $PROJECT & "\" & StringUpper($CliParams[$CLI_COMPANYCODE]) & $PROJECT & "_" & $AUTO_FOLDER_SOURCE, $DIR_EXTENDED)
   If $ResidualFiles[1] > 0 Then
	  LogEvent($LogFile, "Residual file (" & $ResidualFiles[1] & ") found in the local source location " & $AUTO_FOLDER_ROOT & "\" & StringUpper($CliParams[$CLI_COMPANYCODE]) & $PROJECT & "\" & StringUpper($CliParams[$CLI_COMPANYCODE]) & $PROJECT & "_" & $AUTO_FOLDER_SOURCE, True)
	  MessageToAdmin("E;" & @ScriptName & ";" & $CliParams[$CLI_SYSTEM],"Residual file(s) found in the local source. See the attached log file", $SYS_ADMINS, $LogFilePath)
	  $SAP[$SAP_CON].CloseConnection ; Clean up
	  Exit(98)
   EndIf

   ; ******************************************************************
   ; STATE_INITIAL -> STATE_INVALID || STATE_OK || STATE_ALREADYEXISTS
   ;*******************************************************************
   ; Initial state. Each file goes through state changes as it's  being processed
   ; Initial filtering process examines the file type and file size and sends the
   ; file to either STATE_INVALID (bad file type or size limit exceede) or
   ; STATE_OK or STATE_ALREADYEXISTS (residual file)

   ;Skip files larger than the maximum size limit
   If $File[$FILE_SIZE] >= $MAX_SIZEINBYTES Then
	  If $MAX_SIZEINBYTES <> -1 Then
		 LogEvent($LogFile, $File[$FILE_NAME] & " (" & $File[$FILE_PARENTFOLDER] & ") (" & $FileState[$File[$FILE_STATUS]] & ")", False)
		 LogEvent($LogFile, $File[$FILE_NAME] & " (" & $File[$FILE_PARENTFOLDER] & ") exceeding the $MAX_SIZEINBYTES (" & $MAX_SIZEINBYTES & ") File size -> " & $File[$FILE_SIZE], False)
		 $File[$FILE_STATUS] = $STATE_INVALID ; Set state STATE_INVALID (file size)
		 LogEvent($LogFile, $File[$FILE_NAME] & " (" & $File[$FILE_PARENTFOLDER] & ") (" & $FileState[$File[$FILE_STATUS]] & ")", False)
	  EndIf
   ElseIf Not StringRegExp($File[$FILE_NAME],"pdf$") Then
	  LogEvent($LogFile, $File[$FILE_NAME] & " (" & $File[$FILE_PARENTFOLDER] & ") (" & $FileState[$File[$FILE_STATUS]] & ")", False)
	  LogEvent($LogFile, $File[$FILE_NAME] & " is not a valid .pdf file", False)
	  $File[$FILE_STATUS] = $STATE_INVALID ; Set state STATE_INVALID (file type)
	  LogEvent($LogFile, $File[$FILE_NAME] & " (" & $File[$FILE_PARENTFOLDER] & ") (" & $FileState[$File[$FILE_STATUS]] & ")", False)
   EndIf

    ;Something went horribly wrong if the file is still there. If script fails before moving file to the archive dir
   If $CliParams[$CLI_SUBDIRS] And $File[$FILE_STATUS] <> $STATE_INVALID then
	  If FileExists($PathSrc & "\" & $File[$FILE_PARENTFOLDER] & "_" & $File[$FILE_NAME]) Then
		 LogEvent($LogFile, $File[$FILE_NAME] & " (" & $File[$FILE_PARENTFOLDER] & ") (" & $FileState[$File[$FILE_STATUS]] & ")", False)
		 LogEvent($LogFile, "File " & $File[$FILE_NAME] & " (" & $File[$FILE_PARENTFOLDER] & ") already exists in the lolcal source folder. Skipping this file", False)
		 $File[$FILE_STATUS] = $STATE_ALREADYEXISTS ; Set state ALREADY_EXISTS
		 LogEvent($LogFile, $File[$FILE_NAME] & " (" & $File[$FILE_PARENTFOLDER] & ") (" & $FileState[$File[$FILE_STATUS]] & ")", False)
	  Else
		 LogEvent($LogFile, $File[$FILE_NAME] & " (" & $File[$FILE_PARENTFOLDER] & ") (" & $FileState[$File[$FILE_STATUS]] & ")", False)
		 $File[$FILE_STATUS] = $STATE_OK ; Set state STATE_OK
		 LogEvent($LogFile, $File[$FILE_NAME] & " (" & $File[$FILE_PARENTFOLDER] & ") (" & $FileState[$File[$FILE_STATUS]] & ")", False)
	  EndIf
   Elseif Not $CliParams[$CLI_SUBDIRS] And $File[$FILE_STATUS] <> $STATE_INVALID Then
	  If FileExists($PathSrc & "\" & $File[$FILE_NAME]) Then
		 LogEvent($LogFile, $File[$FILE_NAME] & " (" & $File[$FILE_PARENTFOLDER] & ") (" & $FileState[$File[$FILE_STATUS]] & ")", False)
		 LogEvent($LogFile, "File " & $File[$FILE_NAME] & " (" & $File[$FILE_PARENTFOLDER]  & ") already exists in the lolcal source folder. Skipping this file", False)
		 $File[$FILE_STATUS] = $STATE_ALREADYEXISTS ; Set state ALREADY_EXISTS
		 LogEvent($LogFile, $File[$FILE_NAME] & " (" & $File[$FILE_PARENTFOLDER] & ") (" & $FileState[$File[$FILE_STATUS]] & ")", False)
	  Else
		 LogEvent($LogFile, $File[$FILE_NAME] & " (" & $File[$FILE_PARENTFOLDER] & ") (" & $FileState[$File[$FILE_STATUS]] & ")", False)
		 $File[$FILE_STATUS] = $STATE_OK ; Set state STATE_OK
		 LogEvent($LogFile, $File[$FILE_NAME] & " (" & $File[$FILE_PARENTFOLDER] & ") (" & $FileState[$File[$FILE_STATUS]] & ")", False)
	  EndIf
   EndIf


   ; ******************************************************************
   ; STATE_OK -> STATE_COPYBEGIN || STATE_COPIED
   ; ******************************************************************
   ; If the file is in the STATE_OK state, we continue with the file download
   ; where file gets into the STATE_COPYBEGIN and if the copying has been
   ; successfull, then the state is changed to SATE_COPIED

   ;Try copying the file
   If $File[$FILE_STATUS] = $STATE_OK Then
	  $File[$FILE_STATUS] = $STATE_COPYBEGIN
	  LogEvent($LogFile, $File[$FILE_NAME] & " (" & $File[$FILE_PARENTFOLDER] & ") (" & $FileState[$File[$FILE_STATUS]] & ")", False)

	  If $CliParams[$CLI_SUBDIRS] Then
		 If FileCopy($File[$FILE_PATH], $PathSrc & "\" & $File[$FILE_PARENTFOLDER] & "_" & $File[$FILE_NAME]) Then
			$File[$FILE_STATUS] = $STATE_COPIED
			LogEvent($LogFile, $File[$FILE_NAME] & " (" & $File[$FILE_PARENTFOLDER] & ") (" & $FileState[$File[$FILE_STATUS]] & ")", False)
		 EndIf
	  Else
		 If FileCopy($File[$FILE_PATH], $PathSrc & "\" & $File[$FILE_NAME]) Then
			$File[$FILE_STATUS] = $STATE_COPIED
			LogEvent($LogFile, $File[$FILE_NAME] & " (" & $File[$FILE_PARENTFOLDER] & ") (" & $FileState[$File[$FILE_STATUS]] & ")", False)
		 EndIf
	  EndIf
   EndIf

   ; ******************************************************************
   ; STATE_COPIED -> STATE_DRAGDROREADY
   ; ******************************************************************
   If $File[$FILE_STATUS] = $STATE_COPIED Then
	  If IsExplorerOpen($PathSrc, $ExplorerTitle) Then
		 ExplorerSetView($ExplorerTitle)
		 $File[$FILE_STATUS] = $STATE_DRAGDROPREADY
		 LogEvent($LogFile, $File[$FILE_NAME] & " (" & $File[$FILE_PARENTFOLDER] & ") (" & $FileState[$File[$FILE_STATUS]] & ")", False)
	  EndIf
   EndIf



   ; *****************************************************************
   ; STATE_DRAGDROPREADY -> STATE_DRAGDROPBEGIN -> STATE_DRAGDROPPED
   ; *****************************************************************
   ;Drag & Drop
   If $File[$FILE_STATUS] = $STATE_DRAGDROPREADY Then
	  $File[$FILE_STATUS] = $STATE_DRAGDROPBEGIN
	  LogEvent($LogFile, $File[$FILE_NAME] & " (" & $File[$FILE_PARENTFOLDER] & ") (" & $FileState[$File[$FILE_STATUS]] & ")", False)

	  $DropResult = DragDrop($SessionOAWD, $ExplorerWindowTitle)

	  If StringRegExp($DropResult, "Action completed") Then
		 $File[$FILE_STATUS] = $STATE_DRAGDROPPED
		 LogEvent($LogFile, $File[$FILE_NAME] & " (" & $File[$FILE_PARENTFOLDER] & ") (" & $FileState[$File[$FILE_STATUS]] & ")", False)
		 $WinOAWD.SendVKey(0) ; This closes the dragdrop window only if drop was succesfull. If not, window remians opened for amdin inspection
	  Else
		 ; We could recover from this but for the sake of simplicity we won't
		 $File[$FILE_STATUS] = $STATE_DRAGDROPFAILED
		 LogEvent($LogFile, $File[$FILE_NAME] & " (" & $File[$FILE_PARENTFOLDER] & ") (" & $FileState[$File[$FILE_STATUS]] & ")", False)
		 LogEvent($LogFile, $DropResult, True)
		 ; Cleanup
		 FileDelete($LocalSourceFolder)
		 MessageToAdmin("E;" & @ScriptName & ";" & $CliParams[$CLI_SYSTEM], "Drag and drop failed with the following file: " & $File[$FILE_NAME] & " (" & $File[$FILE_PARENTFOLDER] & ")" & @CRLF _
		              & "State: " & $FileState[$File[$FILE_STATUS]] & @CRLF _
					  & "SAP Error: " & $DropResult ,$SYS_ADMINS, $LogFilePath)
		 $SAP[$SAP_CON].CloseConnection
		 While WinExists($ExplorerTitle,"")
			WinKill($ExplorerTitle,"")
		 WEnd
		 Exit(205)
	  EndIf
   EndIf



   ; ***********************************************************************
   ; STATE_DRAGDROPPED -> [ STATE_ARCHIVELOCALBEGIN -> STATE_ARCHIVEDLOCAL ]
   ; ***********************************************************************
   If $CliParams[$CLI_ARCHIVELOCAL] Then
	  If $File[$FILE_STATUS] = $STATE_DRAGDROPPED Then
	     DirCreate($LocalArchiveFolder)
		 $File[$FILE_STATUS] = $STATE_ARCHIVELOCALBEGIN
		 LogEvent($LogFile, $File[$FILE_NAME] & " (" & $File[$FILE_PARENTFOLDER] & ") (" & $FileState[$File[$FILE_STATUS]] & ")", False)

		 If $CliParams[$CLI_SUBDIRS] Then
			Local $__localSourceFileName = $File[$FILE_PARENTFOLDER] & "_" & $File[$FILE_NAME]
		 Else
			Local $__localSourceFileName = $File[$FILE_NAME]
		 EndIf


         If FileCopy($LocalSourceFolder & "\" & $__localSourceFileName, $LocalArchiveFolder & "\" & $__localSourceFileName) Then
			$File[$FILE_STATUS] = $STATE_ARCHIVEDLOCAL
			LogEvent($LogFile, $File[$FILE_NAME] & " (" & $File[$FILE_PARENTFOLDER] & ") (" & $FileState[$File[$FILE_STATUS]] & ")", False)
		 EndIf
	  EndIf
   EndIf

   ; ****************************************************************************************
   ; STATE_ARCHIVEDLOCAL || STATE_DRAGDROPPED -> [ STATE_ARCHIVESPBEGIN -> STATE_ARCHIVEDSP ]
   ; ****************************************************************************************
   If $CliParams[$CLI_ARCHIVESHAREPOINT] Then
	  If $File[$FILE_STATUS] = $STATE_ARCHIVEDLOCAL Or $File[$FILE_STATUS] = $STATE_DRAGDROPPED Then

		 $File[$FILE_STATUS] = $STATE_ARCHIVESPBEGIN
		 LogEvent($LogFile, $File[$FILE_NAME] & " (" & $File[$FILE_PARENTFOLDER] & ") (" & $FileState[$File[$FILE_STATUS]] & ")", False)

		 If Not $CliParams[$CLI_SUBDIRS] Then
			If SPFileUpload($SpAccessToken, $PathSrc & "\" & $File[$FILE_NAME], $CliParams[$CLI_ARCHIVESHAREPOINTURL], $RunTimeStamp) Then
			   $File[$FILE_STATUS] = $STATE_ARCHIVEDSP
			   LogEvent($LogFile, $File[$FILE_NAME] & " (" & $File[$FILE_PARENTFOLDER] & ") (" & $FileState[$File[$FILE_STATUS]] & ")", False)
			EndIf
		 Else
			If SPFileUpload($SpAccessToken, $PathSrc & "\" & $File[$FILE_PARENTFOLDER] & "_" & $File[$FILE_NAME], $CliParams[$CLI_ARCHIVESHAREPOINTURL], $RunTimeStamp) Then
			   $File[$FILE_STATUS] = $STATE_ARCHIVEDSP
			   LogEvent($LogFile, $File[$FILE_NAME] & " (" & $File[$FILE_PARENTFOLDER] & ") (" & $FileState[$File[$FILE_STATUS]] & ")", False)
			EndIf
	     EndIf
	  EndIf
   EndIf

   ; **************************************************************************************************************
   ; STATE_ARCHIVEDSP || STATE_DRAGDROPPED || STATE_ARCHIVEDLOCAL -> STATE_DELETEREMOTEBEGIN -> STATE_DELETEDREMOTE
   ; **************************************************************************************************************
    If $File[$FILE_STATUS] = $STATE_ARCHIVEDSP Or $File[$FILE_STATUS] = $STATE_ARCHIVEDLOCAL Or $File[$FILE_STATUS] = $STATE_DRAGDROPPED Then
	  $File[$FILE_STATUS] = $STATE_DELETEREMOTEBEGIN
	  LogEvent($LogFile, $File[$FILE_NAME] & " (" & $File[$FILE_PARENTFOLDER] & ") (" & $FileState[$File[$FILE_STATUS]] & ")", False)

	  If $DataSource = $SOURCE_NETDRIVE Then
		 If FileDelete($File[$FILE_PATH]) Then
			FileDelete($LocalSourceFolder) ; Also safely delete local copy in the source folder
			$File[$FILE_STATUS] = $STATE_DELETEDREMOTE
			LogEvent($LogFile, $File[$FILE_NAME] & " (" & $File[$FILE_PARENTFOLDER] & ") (" & $FileState[$File[$FILE_STATUS]] & ")", False)
		 EndIf
	  Elseif $DataSource = $SOURCE_SHAREPOINT Then
		 ; Not yet implemented
	  EndIf
    EndIf

   ; ******************************************************************
   ; If the state at this point is not $STATE_DELETEDREMOTE script ends
   ; ******************************************************************
   If $File[$FILE_STATUS] <> $STATE_DELETEDREMOTE Then
	  LogEvent($LogFile,"", True) ; Just close the log file
	  MessageToAdmin("E;" & @ScriptName & ";" & $CliParams[$CLI_SYSTEM], "File " & $File[$FILE_NAME] & " (" & $File[$FILE_PARENTFOLDER] & ") did not reach the final state. Most recent state is: " & @CRLF _
				   & "State: " & $FileState[$File[$FILE_STATUS]], $SYS_ADMINS, $LogFilePath)
	  $SAP[$SAP_CON].CloseConnection
	  While WinExists($ExplorerTitle,"")
		 WinKill($ExplorerTitle,"")
	  WEnd
	  Exit(199)
   EndIf


   LogEvent($LogFile, "-------------------------------------------------------", False)
Next
#EndRegion

#Region Final cleanup
While WinExists($ExplorerTitle,"")
   WinKill($ExplorerTitle,"")
WEnd
FileDelete($LocalSourceFolder)
$SAP[$SAP_CON].CloseConnection
LogEvent($LogFile,"Calling Wdap()", False)
Wdapp(StringUpper($CliParams[$CLI_COMPANYCODE]) & $PROJECT)
LogEvent($LogFile,"DONE", True)
MessageToAdmin("I;" & @ScriptName & ";" & $CliParams[$CLI_SYSTEM],"Processing of (" &  $FilesToProcess.Count & ") files took " & Floor(TimerDiff($Timer) / 1000) & " seconds" ,$SYS_ADMINS, $LogFilePath)
Exit(0)
#EndRegion

#EndRegion


#Region Network share functions

Func NSFolderGetFileCount($_sNetworkPath, $_processSubdirs, ByRef $_dFilesToProcess)
   Local $Fso = ObjCreate("Scripting.FileSystemObject")
   Local $Folder = $Fso.GetFolder($_sNetworkPath)
   Local $Subfolder
   Local $File
   Local $__SubfoldersCount = 0

   ; Subfolders
   If $_processSubdirs And $Folder.Subfolders.Count = 0 Then
	  SetError(105) ; Subdirectories processing requested but no matching subfolders found
	  Return 0
   ElseIf $_processSubdirs And $Folder.Subfolders.Count >= 1 Then
	  For $Subfolder in $Folder.Subfolders
		 If StringRegExp($Subfolder.Name,StringUpper($CliParams[$CLI_COMPANYCODE])) Then
			$__SubfoldersCount = $__SubfoldersCount + 1
			For $File In $Subfolder.Files

			   ;$FILE_PARENTFOLDER = 0, $FILE_PATH, $FILE_NAME, $FILE_SIZE, $FILE_SOURCETYPE, $FILE_STATUS
			   Local $__filePropertiesArray[6] = [$File.ParentFolder.Name, $File.Path, $File.Name, $File.Size, $SOURCE_NETDRIVE, $STATE_INTIAL]
			    $_dFilesToProcess.Add($Subfolder.Name & "_" & $File.Name, $__filePropertiesArray)

			Next
		 EndIf
	  Next

	  Return $_dFilesToProcess.Count
   EndIf

   ;No subfolders
   If $Folder.Files.Count > 0 Then
	  For $File In $Folder.Files

		 ;$FILE_PARENTFOLDER = 0, $FILE_PATH, $FILE_NAME, $FILE_SIZE, $FILE_SOURCETYPE, $FILE_STATUS
		 Local $__filePropertiesArray[6] = [$File.ParentFolder.Name, $File.Path, $File.Name, $File.Size, $SOURCE_NETDRIVE, $STATE_INTIAL]
		 $_dFilesToProcess.Add($File.Name, $__filePropertiesArray)

	  Next

	  Return $_dFilesToProcess.Count
   Else
	  Return 0
   EndIf
EndFunc
#EndRegion

#Region Sharepoint functions

Func SPFolderGetFileCount(ByRef $_oHTTP, $_sSiteUrl, $_sSecurityToken, $_processSubdirs)
     Local $__XmlFiles = ObjCreate("MSXML2.DOMDocument")
	 Local $__XmlFolders = ObjCreate("MSXML2.DOMDocument")
	 Local $__Node = Null
	 Local $__FileCount = 0
	 Local $__SubfoldersCount = 0

	 If StringRight($_sSiteUrl,1) == "/" Then
	    $_sSiteUrl = StringTrimRight($_sSiteUrl,1)
	 EndIf

	 Local $__SharepointDir = StringRegExp($_sSiteUrl,"\/([^\/]+)\/?$", $STR_REGEXPARRAYMATCH)[0]

	 $_sSiteUrl = StringTrimRight($_sSiteUrl, StringLen($__SharepointDir))
     $_sSiteUrl = $_sSiteUrl & "_api/web/GetFolderByServerRelativeUrl('" & $__SharepointDir & "')"


	With $_oHTTP
		.open("GET", $_sSiteUrl, False)
		.setRequestHeader("Authorization", "Bearer " & $_sSecurityToken)
		.setRequestHeader("Accept", "application/atom+xml")
		.send()
	EndWith


	If $_oHTTP.status == 404 Then
		; Folder not found and @error is set
		SetError(305) ; Folder not found
		Return 0
	EndIf

	; Process files in the root directory
    If Not $_processSubdirs Then

	  With $_oHTTP
		.open("GET", $_sSiteUrl & "/Files", False)
		.setRequestHeader("Authorization", "Bearer " & $_sSecurityToken)
		.setRequestHeader("Accept", "application/atom+xml")
		.send()
	  EndWith

	  If $_oHTTP.status <> 200 Then
		SetError(306) ;Cant obtain files
		Return 0
	  EndIf

	  $__XmlFiles.LoadXML($_oHTTP.responseText)
	  $__FileCount = $__XmlFiles.selectNodes("//id").length

	  Return $__FileCount

   ; Process subdirectories
   Else

	  With $_oHTTP
	    .open("GET", $_sSiteUrl & "/Folders", False)
	    .setRequestHeader("Authorization", "Bearer " & $_sSecurityToken)
		.setRequestHeader("Accept", "application/atom+xml")
		.send()
	 EndWith

	 If $_oHTTP.status <> 200 Then
		SetError(307) ;Can't obtain subfolders
		Return 0
	 EndIf

	 $__XmlFolders.LoadXML($_oHTTP.responseText)

	  ;//m:properties/d:Name
	  ;//m:properties/d:Name[contains(text(),""SI01"")]      <- Doesn't work :(
     For $__Node in $__XmlFolders.selectNodes("//m:properties/d:Name")

		If StringRegExp($__Node.text,StringUpper($CliParams[$CLI_COMPANYCODE])) Then
		   $__SubfoldersCount = $__SubfoldersCount + 1
		   ConsoleWrite("Counting files in sub folder: " & $__Node.text & @CRLF)
	    EndIf

	  Next

	  If $__SubfoldersCount = 0 And $_processSubdirs Then
		 SetError(308) ; Subfolders processing requested but no subfolders found
		 Return(0)
	  Endif
   EndIf


 EndFunc


Func SPGetAccessToken(ByRef $_oHTTP, $_sTenantDomainName, $_SiteURL, $_AppID, $_AppSecret)

    Local $__sResponseHeader
	Local $__aStringSplit1
	Local $__aStringSplit2
	Local $__sToken
	Local $__sHttpBody
	Local $__TenantID
	Local $__ClientID


	If StringRight($_SiteURL,1) == "/" Then
		$_SiteURL = $_SiteURL & "_vti_bin/client.svc"
	Else
		$_SiteURL = $_SiteURL & "/_vti_bin/client.svc"
	EndIf

	With $_oHTTP
		.open("GET", $_SiteURL, False)
		.setRequestHeader("Authorization", "Bearer")
		.send()
	EndWith


	If $_oHTTP.Status == 401 Then ; 401 is expected at this stage
		$__sResponseHeader = $_oHTTP.getResponseHeader("WWW-Authenticate")

		If Not StringRegExp($__sResponseHeader,"realm=""[a-f0-9]{8}-[a-f0-9]{4}-[a-f0-9]{4}-[a-f0-9]{4}-[a-f0-9]{12}") Then
		   Return SetError(302) ; Error getting RealmID
	    EndIf

		If Not StringRegExp($__sResponseHeader,"client_id=""[a-f0-9]{8}-[a-f0-9]{4}-[a-f0-9]{4}-[a-f0-9]{4}-[a-f0-9]{12}") Then
		   Return SetError(303) ; Error getting ClientID
	    EndIf
	 Else
		Return SetError(301) ; Error getting Tenant/RealmID
     EndIf

   $__TenantID = StringRegExpReplace(StringRegExp($__sResponseHeader,"realm=""[a-f0-9]{8}-[a-f0-9]{4}-[a-f0-9]{4}-[a-f0-9]{4}-[a-f0-9]{12}",$STR_REGEXPARRAYMATCH)[0],"realm=""","")
   $__ClientID = StringRegExpReplace(StringRegExp($__sResponseHeader,"client_id=""[a-f0-9]{8}-[a-f0-9]{4}-[a-f0-9]{4}-[a-f0-9]{4}-[a-f0-9]{12}",$STR_REGEXPARRAYMATCH)[0],"client_id=""","")

	If StringLeft($_sTenantDomainName,1) <> "/" Then
		$_sTenantDomainName = "/" & $_sTenantDomainName
	EndIf

	$__sHttpBody = "grant_type=client_credentials&client_id=" & $_AppID & "@" & $__TenantID & "&client_secret=" & $_AppSecret & "&resource=" & $__ClientID & $_sTenantDomainName & "@" & $__TenantID

	With $_oHTTP
		.open("POST", "https://accounts.accesscontrol.windows.net/" & $__TenantID & "/tokens/OAuth/2", False)
		.setRequestHeader("Content-Type", "application/x-www-form-urlencoded")
		.setRequestHeader("Content-Length", StringLen($__sHttpBody))
		.send($__sHttpBody)
	 EndWith


	If $_oHTTP.status == 200 Then
		Return StringRegExp($_oHTTP.ResponseText,"\""access_token\"":\""(.*)\""",$STR_REGEXPARRAYFULLMATCH)[1]
	 Else
		Return SetError(304)
    EndIf

EndFunc


Func SPFileUpload($token, $source, $destination, $prefix)
   If StringRight($destination,1) = "/" Then
	  $destination = StringLeft($destination, StringLen($destination) - 1)
   EndIf

   Local $http = ObjCreate("winhttp.winhttprequest.5.1")
   $spfolder = StringRegExp($destination, "([^\/]+)(?=[^\/]*\/?$)", $STR_REGEXPARRAYMATCH)[0]
   $destination = StringRegExpReplace($destination, "([^\/]+)(?=[^\/]*\/?$)", "")
   $filename = $prefix & "_" & StringRegExp($source, "([^\\]+)(?=[^\\]*\/?$)", $STR_REGEXPARRAYMATCH)[0]

   Local $hFile = FileOpen($source, $FO_BINARY)
   Local $buffer = FileRead($hFile)
   Local $bufferLen = FileGetSize($source)
   FileClose($hFile)

   With $http
		.open("POST",$destination & "_api/Web/GetFolderByServerRelativeUrl('" & $spfolder & "')/Files/add(overwrite=false, url='" & StringRegExpReplace($filename,"'","%27%27") & "')", False)
		.setRequestHeader("accept", "application/json;odata=verbose")
		.setRequestHeader("Authorization", "Bearer " & $token)
		.setRequestHeader("Content-Length", $bufferLen)
		.send($buffer)
   EndWith

   If $http.Status <> 200 Then
	  LogEvent($LogFile, $File[$FILE_NAME] & " (" & $File[$FILE_PARENTFOLDER] & ") HTTP STATUS: " & $http.Status , False)
	  LogEvent($LogFile, $File[$FILE_NAME] & " (" & $File[$FILE_PARENTFOLDER] & ") HTTP RESPONSE: " & $http.responseText , False)
	  Return False
   EndIf

   Return True
EndFunc
#EndRegion

#Region Logging
Func LogEvent($_LogFile, $_Message, $_CloseFile)
   FileWriteLine($_LogFile, "[" & @MDAY & "." & @MON & "." & @YEAR & " - " & @HOUR & ":" & @MIN & ":" & @SEC & "] " & @ScriptName & " >> " & $_Message)

   If $_CloseFile Then
	  FileClose($_LogFile)
   EndIf
EndFunc
#EndRegion

#Region Credentials
Func GetCredentials($_resourceName)
   Local Enum $credentialUser = 0, $credentialPassword, $credentialDomain, $credentialHost ; 0,1,2,3
   Local $__conectionString = "Provider=Microsoft.ACE.OLEDB.12.0;WSS;IMEX=1;RetrieveIds=Yes;DATABASE=https://volvogroup.sharepoint.com/sites/unit-rc-sk-bs-it/CREDENTIALS;LIST=CREDENTIALS;"
   Local $__adodbConnection = ObjCreate("Adodb.Connection")
   Local $__adodbRecordset  = ObjCreate("Adodb.Recordset")
   Local $__credentials[4]

   $__adodbConnection.ConnectionString = $__conectionString
   $__adodbConnection.Open


   $__adodbRecordset.Open("SELECT Host,Username,Password FROM [CREDENTIALS] WHERE Title='" & $_resourceName & "';", $__adodbConnection, 3, 3)

   If $__adodbRecordset.EOF Or $__adodbRecordset.BOF Then
	  Return 0
   EndIf

   $__adodbRecordset.MoveFirst

   $__credentials[$credentialUser] = $__adodbRecordset.Fields("Username").Value
   $__credentials[$credentialPassword] = $__adodbRecordset.Fields("Password").Value
   $__credentials[$credentialDomain] = StringRegExp($__adodbRecordset.Fields("Host").Value, "^(?:https?:\/\/)?(?:[^@\n]+@)?(?:www\.)?([^:\/\n?]+)", $STR_REGEXPARRAYMATCH)[0]
   $__credentials[$credentialHost] = $__adodbRecordset.Fields("Host").Value

   $__adodbRecordset.Close
   $__adodbConnection.Close

   Return $__credentials
EndFunc
#EndRegion

#Region SAP functions
;---------------------- SAP Functions ---------------------------
Func SAPLaunchSAPLogon(ByRef $_pidStorage)

   Local $_oWMI = ObjGet("winmgmts:\\.\root\cimv2")
   Local $_colProc = $_oWMI.ExecQuery("Select Name, ProcessId From Win32_Process Where Name Like '%saplogon%'")

   If IsObj($_colProc) and $_colProc.count > 0 Then
		For $_proc in $_colProc
			If StringInStr($_proc.Name,"saplogon",$STR_NOCASESENSE) > 0 Then ; Substring found
				$_pidStorage = $_proc.ProcessId
				Return True
				ExitLoop
			EndIf
		Next
   EndIf

   Local $__pid
   Run("C:\Program Files (x86)\SAP\FrontEnd\SapGui\saplogon.exe","",@SW_MINIMIZE)
   $__pid = ProcessWait("saplogon.exe", 10)

   If $__pid == 0 Then
	  Return False ; SAPLogon couldn't be started
   Else
	  $_pidStorage = $__pid
	  Return True
   EndIf

   WinSetState("SAP Logon 750","", @SW_MINIMIZE )
EndFunc

Func SAPSGetSystemDescription($_sSystemName)
	Local $_oXML = ObjCreate("MSXML2.DOMDocument")
	Local $_colNodes,$_node
	$_oXML.load($SAP_LOCAL_LANDSCAPE_PATH)
	$_colNodes = $_oXML.GetElementsByTagName("Service")
	For $_node in $_colNodes
		If(StringLeft(StringLower($_node.attributes.getNamedItem("name").text),3)) == StringLower($_sSystemName) Then
			Return $_node.attributes.getNamedItem("name").text
			ExitLoop
		EndIf
	Next
	Return Null
EndFunc

Func SAPOpenConnection($_SystemName,$_ClientName,$_UserName)
   ; Wait a bit so that sap logon is fully initialized otherwise we get 0x800401E4
   Sleep(3000)
   Local $__SapGui =  ObjGet("SAPGUI")
   Local $__GuiApplication
   Local $__GuiSession
   Local $__GuiConnection

	While Not IsObj($__SapGui)
	   $__SapGui = ObjGet("SAPGUI")
    WEnd

	$__GuiApplication = $__SapGui.GetScriptingEngine

	If Not IsObj($__GuiApplication) Then
	   Return 201 ; Can't obtain Scripting Engine
    EndIf

	$__GuiConnection = $__GuiApplication.OpenConnection(SAPSGetSystemDescription($_SystemName),True,False)

	If Not IsObj($__GuiConnection) Then
	   Return Null
    Endif

    $__GuiSession = $__GuiConnection.Children.Item(0) ; 1st session of the connection

	SAPKillPopups($__GuiSession)

	If StringInStr($__GuiSession.ActiveWindow.FindByName("sbar", "GuiStatusbar").text, "Enter a valid SAP user or choose one from the list") > 0 Then
	  $__GuiSession.ActiveWindow.findById("usr/txtRSYST-MANDT").text = $_ClientName
	  $__GuiSession.ActiveWindow.findById("usr/txtRSYST-BNAME").text = $_UserName
	  $__GuiSession.ActiveWindow.findById("usr/txtRSYST-LANGU").text = "EN"
	  $__GuiSession.ActiveWindow.sendvkey(0)
	  SAPKillPopups($__GuiSession)
    EndIf

    $SAP[$SAP_APP] = $__GuiApplication
	$SAP[$SAP_CON] = $__GuiConnection
	$SAP[$SAP_SES] = $__GuiSession

    Return $__GuiConnection

EndFunc

Func SAPKillPopups($_GuiSession)
	While $_GuiSession.Children.Count > 1
		If StringInStr($_GuiSession.ActiveWindow.Text, "System Message") > 0 Then
			$_GuiSession.ActiveWindow.sendVKey(12)
		ElseIf StringInStr($_GuiSession.ActiveWindow.Text, "Information") > 0 And StringInStr($_GuiSession.ActiveWindow.PopupDialogText, "Exchange rate adjusted to system settings") > 0 Then
			$_GuiSession.ActiveWindow.sendVKey(0)
		ElseIf StringInStr($_GuiSession.ActiveWindow.Text, "Copyright") > 0 Then
			$_GuiSession.ActiveWindow.sendVKey(0)
		ElseIf StringInStr($_GuiSession.ActiveWindow.Text, "License Information for Multiple Logon") > 0 Then
			$_GuiSession.findById("wnd[1]/usr/radMULTI_LOGON_OPT2").select
			$_GuiSession.ActiveWindow.sendVKey(0)
		;ElseIF   'Insert next type of popup windows which you want to kill
		Else
			ExitLoop
		EndIf
	Wend
EndFunc


 #EndRegion

#Region Wdapp
 Func Wdapp($sProjectName)
   Local $oXML = ObjCreate("MSXML2.DOMDocument")
   Local $oHTTP = ObjCreate("winhttp.winhttprequest.5.1")
   Local $oSYSINFO = ObjCreate("ADSystemInfo")
   Local $oUSER = ObjGet("LDAP://" & $oSYSINFO.UserName)
   Local $oNET = ObjCreate("Wscript.Network")

   With $oHTTP
	  .open("GET", "https://volvogroup.sharepoint.com/sites/unit-rc-sk-bs-it/_api/web/lists/getbytitle('WDAPP')/items?$select=Title&$filter=(Title eq '" & $sProjectName & "')", False)
	  .setRequestHeader("Authorization", "Bearer " & $SpAccessToken)
	  .setRequestHeader("Accept", "application/atom+xml;odata=verbose")
	  .send()
   EndWith

	;Patch record
	$oXML.loadXML($oHTTP.responseText)
	Local $url = $oXML.selectSingleNode("//feed").attributes.getNamedItem("xml:base").text
	$url = $url & $oXML.selectSingleNode("//entry/link[@rel=""edit""]").attributes.getNamedItem("href").text

	With $oHTTP
		.open("PATCH", $url, False)
		.setRequestHeader("Accept","application/json;odata=verbose")
		.setRequestHeader("Content-Type","application/json")
		.setRequestHeader("Authorization","Bearer " & $SpAccessToken)
		.setRequestHeader("If-Match","*")
		.send("{""ComputerName"":""" & $oNET.ComputerName & """,""UserName"":""" & $oUSER.displayName & """,""UserID"":""" & $oUSER.sAMAccountName & """}")
	EndWith
 EndFunc

 #EndRegion


#Region IsWorkstationLocked()
Func IsWorkstationLocked()
  Local Const $WTS_CURRENT_SERVER_HANDLE = 0
  Local Const $WTS_CURRENT_SESSION = -1
  Local Const $WTS_SESSION_INFO_EX = 25

  Local $hWtsapi32dll = DllOpen("Wtsapi32.dll")
  Local $result = DllCall($hWtsapi32dll, "int", "WTSQuerySessionInformation", "int", $WTS_CURRENT_SERVER_HANDLE, "int", $WTS_CURRENT_SESSION, "int", $WTS_SESSION_INFO_EX, "ptr*", 0, "dword*", 0)
  If @error Or Not $result[0] Then Return SetError(1, 0, False)

  Local $buffer_ptr = $result[4], $buffer_size = $result[5]
  Local $buffer = DllStructCreate("uint64 SessionId;uint64 SessionState;int SessionFlags;byte[" & $buffer_size - 20 & "]", $buffer_ptr)
  Local $isLocked = DllStructGetData($buffer, "SessionFlags")

  $buffer = 0
  DllCall($hWtsapi32dll, "int", "WTSFreeMemory", "ptr", $buffer_ptr)
  DllClose($hWtsapi32dll)

  Return $isLocked
EndFunc   ;==>IsWorkstationLocked
#EndRegion

#Region MessageToAdmin()
Func MessageToAdmin($_sSubject, $_sMessage, $_sAdmins, $_logfilePath)
    Local $__mail = ObjCreate("CDO.Message")
	Local $__sysinfo = ObjCreate("ADSystemInfo")
    Local $__user = ObjGet("LDAP://" & $__sysinfo.UserName)

	With $__mail
		.From = $__user.Mail
		.To = $_sAdmins
		.Subject = $_sSubject
	    .AddAttachment($_logfilePath)
		.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/sendusing") = 2
		.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpserver") = "mailgot.it.volvo.net"
		.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpserverport") = 25
		.HTMLBody = $_sMessage
		.Configuration.Fields.Item("urn:schemas:mailheader:X-MSMail-Priority") = "High"
		.Configuration.Fields.Item("urn:schemas:httpmail:importance") = 2
		.Configuration.Fields.Item("urn:schemas:mailheader:X-Priority") = 2
		.Configuration.Fields.Update()
		.Send()
	EndWith


EndFunc
#EndRegion

#Region explorer functions
Func IsExplorerOpen($_path, $_title)

   WinGetPos($_title)
   If @error Then
	  If OpenWindowsExplorer($_path) Then
		 Return True
	  Else
		 Return False
	  EndIf
   Else
	  Return True
   EndIf

EndFunc

Func ExplorerSetView($title)
    Local $__aWindowPosition
    Local $__aControlPosition

	$__aWindowPosition = WinGetPos($title,"")
	$__aControlPosition = ControlGetPos($title,"","DirectUIHWND2")

    MouseMove($__aWindowPosition[0] + $__aControlPosition[0] + Floor($__aControlPosition[2] / 2), $__aWindowPosition[1] + $__aControlPosition[1] + Floor($__aControlPosition[3] / 2))
	SendKeepActive(WinGetHandle($title,""))
	Sleep(500)
	Send("{CTRLDOWN}")
	MouseWheel($MOUSE_WHEEL_UP, 30)
	Send("{CTRLUP}")
	SendKeepActive("")
 EndFunc

Func OpenWindowsExplorer($_path)
    Local $__pid = ShellExecute("explorer.exe",$_path) ; Open explorer and return PID to the caller
	If $__pid == 0 Then
       Return False
	EndIf

	Return True
EndFunc

Func ExplorerRefreshWindow($_hwndWindow, $_nDelayms)
	SendKeepActive($_hwndWindow)
	Send("{F5}")       ; F5
	SendKeepActive("")
	Sleep($_nDelayms)
EndFunc
#EndRegion

#Region DragDrop
Func DragDrop($SapSession, $ExplorerWindow)
    Local $__aMSize ; Main area position w/o navigation pane
	Local $__aESize ; Window position including navigation pane
	Local $__aASize ; SAP Archive from Frontend window position
	Local $__win = $SapSession.Children.Item(1)
	Local $__result

    $__aMSize = ControlGetPos($ExplorerWindow,"","DirectUIHWND2")
	$__aESize = WinGetPos($ExplorerWindow,"")

   ;Reopen archive window and position it at 0,0
	If Not StringRegExp($__win, "Storing for subsequent entry") Then
	   $__win.sendVKey(9)
    EndIf

    WinActivate($__win.Text,"")
    WinMove($__win.Text,"",0,0)
	$__aASize = WinGetPos($__win.Text)



   ;Position explorer window right next to sap archive window
   WinActivate($ExplorerWindow,"")
   WinMove($ExplorerWindow,"", $__aAsize[2] + 10, 0, 700, 450)

   $__aMSize = ControlGetPos($ExplorerWindow,"","DirectUIHWND2")
   $__aESize = WinGetPos($ExplorerWindow,"")

   MouseClickDrag($MOUSE_CLICK_LEFT, $__aESize[0] + $__aMSize[0] + Floor($__aMSize[2] / 2), $__aESize[1] + $__aMSize[1] + Floor($__aMSize[3] / 2), $__aASize[0] + Floor(0.75 * $__aASize[2]), $__aASize[1] + Floor(0.5 * $__aASize[3]), 10)

   While $SapSession.Busy
	  Sleep(1000)
   WEnd


   ; usually it's Error in HTTP Access: IF_HTTP_CLIENT->RECEIVE 1
   If $SapSession.Children.Count > 2 Then

	  If StringRegExp(StringUpper($SapSession.ActiveWindow.Text), "ERROR") Then
		 $__result = $SapSession.ActiveWindow.findById("usr/txtMESSTXT1").Text
		 Return $__result
	  Else
		 $__result = $SapSession.ActiveWindow.Text
		 Return $__result
	  EndIf
   EndIf

   If Not StringRegExp($SapSession.Children.Item(0).findById("sbar/pane[0]").Text,"Action completed") Then
	  $__result = $SapSession.Children.Item(0).findById("sbar/pane[0]").Text
	  Return  $__result; Some error or undefined
   Else
	  $__result = $SapSession.Children.Item(0).findById("sbar/pane[0]").Text
	  Return  $__result; Some error or undefined
   EndIf

EndFunc


#EndRegion


; This is a custom error handler
Func ErrFunc($oError)
   ConsoleWrite("Com error")
   MessageToAdmin("W;" & @ScriptName & ";" & @YEAR & "-" & @MON & "-" & @MDAY & ";" & @HOUR & ":" & @MIN & ":" & @SEC & ";" & @UserName & ";" & @ComputerName & ";" & $CliParams[$CLI_SYSTEM],"We intercepted a COM Error ! Number: 0x" & Hex($oError.number, 8) & " Description: " & $oError.windescription & " At line: " & $oError.scriptline, $SYS_ADMINS, Null)
   Exit(999)

EndFunc   ;==>ErrFunc

Func URLEncode($urlText)
    $url = ""
    For $i = 1 To StringLen($urlText)
        $acode = Asc(StringMid($urlText, $i, 1))
        Select
            Case ($acode >= 48 And $acode <= 57) Or _
                    ($acode >= 65 And $acode <= 90) Or _
                    ($acode >= 97 And $acode <= 122)
                $url = $url & StringMid($urlText, $i, 1)
            Case $acode = 32
                $url = $url & "+"
            Case Else
                $url = $url & "%" & Hex($acode, 2)
        EndSelect
    Next
    Return $url
EndFunc   ;==>URLEncode
