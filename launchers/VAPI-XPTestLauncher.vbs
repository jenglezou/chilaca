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
'sVBSFrameworkDir = "c:\chilaca" 'Set in VAPI-XP test in ALM

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
globRunMode = VAPI_XP_TEST
'msgbox "here"
sFilePath = oFS.GetAbsolutePathName(sVBSFrameworkDir & "\clsVBSFramework.vbs")

If oFS.FileExists(sFilePath) Then
	Set oFile = oFS.OpenTextFile(sFilePath, 1, False)
	sScript = oFile.ReadAll
	oFile.Close
	Set oFile = Nothing
	ExecuteGlobal sScript 'load vbs to global scope
	Set oVBSFramework = New clsVBSFramework
	Call oVBSFramework.Main(oVBSFramework.GetTestSpreadsheetFromQC())
Else
	TDHelper.AddStepToRun "File not found", "File in path: " & sFilePath & "does not exist",,,"Failed"
End If
		
Set oFS = Nothing

Set oVBSFramework = Nothing
Set globCurrentTSTest =  Nothing
Set globCurrentTestSet = Nothing
Set globCurrentRun  = Nothing
Set globTDConnection  = Nothing	

SetLocale iOriginalLocale

