'Option Explicit

'definition of global objects
Dim globCurrentTSTest     ' current TSTest instead of QCutil.CurrentTestSetTest
Dim globCurrentTestSet    ' current TestSet instead of QCutil.CurrentTestSet
Dim globCurrentRun        ' current Run instead of QCUtil.CurrentRun
Dim globTDConnection      ' current connection to ALM, instead of QCUtil.QCConnection

Dim oVBSFramework         'The VBSFramework object.
Dim sVBSFrameworkDir      'Folder containing the VBSFramework code and VBS files for application objects
Dim iOriginalLocale
Dim oFS, oFile
Dim sScript,sFilePath
'sVBSFrameworkDir = "c:\chilaca" 'set in uft script

Dim arrArgs ' inputs to launcher

Dim globRunMode ' global variable indicates runMode
Dim globRetVal ' return value of framework run
Dim globResultPath '  path where result spreadsheet will be stored
Dim globDataPath '  path where are data spreadsheets stored
Dim globTestStart 'time when spreadsheet started execution
Dim globeDebug: globDebug = False ' if debug mode is on/of
Dim iDuration '  duration of test execution

'Constants
Const QTP_TEST = 1
Const VAPI_XP_TEST = 2
Const CMD_TEST = 3 
Const QTP_LOCAL_TEST = 4

Set oFS = CreateObject("Scripting.FileSystemObject")

'Set Local Spreadsheet
'sQtpDataPath = sVBSFrameworkDir & "\data\"
'sLocalSpreadSheetPath = sVBSFrameworkDir & "\tests\DigitalStrategyTestAutomation.xls"
'msgbox sLocalSpreadSheetPath
'set to default 
globRunMode = -1

iOriginalLocale = SetLocale("en-gb")  'Get the test script location
'sVBSFrameworkDir = oFS.GetAbsolutePathName(".") 'Set the location of the framework code files
'msgbox sVBSFrameworkDir

' default result path where will be Run result
globResultPath = oFS.GetAbsolutePathName(sVBSFrameworkDir & "\results")
' setting data path for spreadsheets			
globDataPath = oFS.GetAbsolutePathName(sVBSFrameworkDir & "\data\")
'Run mode decisor

Set globCurrentTSTest =  Nothing
Set globCurrentTestSet = Nothing
Set globCurrentRun  = Nothing
Set globTDConnection  = Nothing	
globRunMode = QTP_LOCAL_TEST
If IsEmpty(sLocalSpreadSheetPath) Then
	MsgBox("ERROR: Not local spreadsheet defined")
End If


'msgbox "here"
sFilePath = oFS.GetAbsolutePathName(sVBSFrameworkDir & "\clsVBSFramework.vbs")

'Load clsVBSFramework class to global scope according to test type and execute spreadsheet

If oFS.FileExists(sFilePath) Then

	ExecuteFile sFilePath  'load vbs to global scope
	
	If IsEmpty(bRunner) Then
	'by execute_local spreadhseet
	'glob result path is set by default by using execute local spreadsheet
		If Not IsEmpty(globDataPath) Then
			'globDataPath = sQtpDataPath	
		End If			
		Set oVBSFramework = New clsVBSFramework			
		oVBSFramework.Main(sLocalSpreadSheetPath)
	Else
		'by runner:	
		'if not set, will be used default
			
		If Not IsEmpty(sQtpResultPath) Then
			'globResultPath =  sQtpResultPath & "\"
		End If
		
		If Not IsEmpty(globDataPath) Then	
			'globDataPath = sQtpDataPath
		End If	
		Set oVBSFramework = New clsVBSFramework			
		oVBSFramework.Main(sLocalSpreadSheetPath)
	End If
								
Else		    		   
    Reporter.ReportEvent 1, "File not found", "File in path: " & sVBSFrameworkDir & "\clsVBSFramework.vbs " & " does not exist"
End If

Set oFS = Nothing

'Set oVBSFramework = Nothing
'Set globCurrentTSTest =  Nothing
'Set globCurrentTestSet = Nothing
'Set globCurrentRun  = Nothing
'Set globTDConnection  = Nothing	

SetLocale iOriginalLocale

'return value of run - form cmd-runner reason
TestArgs("ExecTime") = globTestStart
TestArgs("Duration") = iDuration	
