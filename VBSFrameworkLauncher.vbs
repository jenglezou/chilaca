'Option Explicit
'definition of global objects
Dim globCurrentTSTest     ' current TSTest instad of QCutil.CurrentTestSetTest
Dim globCurrentTestSet    ' current TestSet instead of QCutil.CurrentTestSet
Dim globCurrentRun        ' current Run instead of QCUtil.CurrentRun
Dim globTDConnection      ' current connection to ALM, instad of QCUtil.QCConnection

Dim oVBSFramework         'The VBSFramework object.
Dim sVBSFrameworkDir      'Folder containing the VBSFramework code and VBS files for application objects
Dim iOriginalLocale
Dim oFS, oFile
Dim sScript,sFilePath


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

'Dim PATH_RESOURCES : PATH_RESOURCES = "C:\GitHubProjects\chilaca\temp" '".\temp\"
'Dim PATH_TESTS : PATH_TESTS =  "C:\GitHubProjects\chilaca\tests\"
'Dim PATH_HOSTLOGFILE : PATH_HOSTLOGFILE = "C:\GitHubProjects\chilaca\temp\trace.log"
Const PATH_RESOURCES = "C:\GitHubProjects\chilaca\temp\" '".\temp\"
Const PATH_TESTS =  "C:\GitHubProjects\chilaca\tests\"
Const PATH_HOSTLOGFILE = "C:\GitHubProjects\chilaca\temp\trace.log"

Set oFS = CreateObject("Scripting.FileSystemObject")

'set to default 
globRunMode = -1

iOriginalLocale = SetLocale("en-gb")  'Get the test script location
'sVBSFrameworkDir = oFS.GetAbsolutePathName(".") 'Set the location of the framwework code files
sVBSFrameworkDir = "C:\GitHubProjects\chilaca" 'oFS.GetAbsolutePathName(".") 'Set the location of the framwework code files
' default result path where will be Run result
globResultPath = oFS.GetAbsolutePathName(sVBSFrameworkDir & "\results")
' setting data path for spreadsheets			
globDataPath = oFS.GetAbsolutePathName(sVBSFrameworkDir & "\data")
'Run mode decisor

If Not IsEmpty(TDHelper) Then
'this is asigned when is VAPI-XP test ran from ALM
	globRunMode = VAPI_XP_TEST
	
ElseIf IsObject(Description) Then	
	
	If IsEmpty(QCUtil) Then
		'this is asigned when local spreadsheat from QTP run
		Set globCurrentTSTest =  Nothing
		Set globCurrentTestSet = Nothing
		Set globCurrentRun  = Nothing
		Set globTDConnection  = Nothing	
		globRunMode = QTP_LOCAL_TEST
		If IsEmpty(sLocalSpreadSheetPath) Then
			MsgBox("ERROR: Not local spreadsheet defined")
		End If
	Else	
		If QCUtil.IsConnected Then
		
			If QCUtil.CurrentRun Is Nothing Then
				'this is asigned when local spreadsheat from QTP run
				Set globCurrentTSTest =  Nothing				
				Set globCurrentTestSet = Nothing
				Set globCurrentRun  = Nothing
				Set globTDConnection  = Nothing	
				globRunMode = QTP_LOCAL_TEST
				If IsEmpty(sLocalSpreadSheetPath) Then
					MsgBox("ERROR: Not local spreadsheet defined")
				End If
			Else
			'this is asigned when is QUICK_TEST test ran from ALM			
				Set globCurrentTSTest =  QCutil.CurrentTestSetTest
				Set globCurrentTestSet = QCutil.CurrentTestSet
				Set globCurrentRun  = QCUtil.CurrentRun
				Set globTDConnection  = QCUtil.QCConnection
				globRunMode = QTP_TEST			
			End If
		Else
			'this is asigned when local spreadsheat from QTP run		
			Set globCurrentTSTest =  Nothing
			Set globCurrentTestSet = Nothing
			Set globCurrentRun  = Nothing
			Set globTDConnection  = Nothing	
			globRunMode = QTP_LOCAL_TEST
			If IsEmpty(sLocalSpreadSheetPath) Then
				MsgBox("ERROR: Not local spreadsheet defined")
			End If	
		End If	
	End if		
Else 
	'this is asigned when framework is run froM command prompt
	Set globCurrentTSTest =  Nothing
	Set globCurrentTestSet = Nothing
	Set globCurrentRun  = Nothing
	Set globTDConnection  = Nothing	
	globRunMode = CMD_TEST
	'vvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvv
	'for local execution
	'change variable sLocalSpreadSheetPath to yours local spreadsheet
	'vvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvv
	'sLocalSpreadSheetPath = "c:\VBSFramework\tests\test\HelloWorldTest.xls"
	'---------------------------------

End If


sFilePath = oFS.GetAbsolutePathName(sVBSFrameworkDir & "\clsVBSFramework.vbs")

'MsgBox "Run mode: " & globRunMode & ", Framework File: " & sFilePath

'Load clsVBSFramework class to global scope according to test type and execute spreadsheet
Select Case globRunMode

	Case QTP_TEST	
	
		If oFS.FileExists(sFilePath) Then
			ExecuteFile sFilePath  'load vbs to global scope
			Set oVBSFramework = New clsVBSFramework
			oVBSFramework.Main oVBSFramework.GetTestSpreadsheetFromQC()
			'oVBSFramework.Main(sLocalSpreadSheetPath) 
		Else		    		   
		    Reporter.ReportEvent 1, "File not found", "File in path: " & sVBSFrameworkDir & "\clsVBSFramework.vbs " & " does not exist"
		End If
		
	Case QTP_LOCAL_TEST
	
		If oFS.FileExists(sFilePath) Then
		
			ExecuteFile sFilePath  'load vbs to global scope
			
			If IsEmpty(bRunner) Then
			'by execute_local spreadhseet
			'glob result path is set by default by using execute local spreadsheet
				If Not IsEmpty(globDataPath) Then
					globDataPath = sQtpDataPath	
				End If			
				Set oVBSFramework = New clsVBSFramework			
				oVBSFramework.Main(sLocalSpreadSheetPath)
			Else
				'by runner:	
				'if not set, will be used default
					
				If Not IsEmpty(sQtpResultPath) Then
					globResultPath =  sQtpResultPath & "\"
				End If
				
				If Not IsEmpty(globDataPath) Then	
					globDataPath = sQtpDataPath
				End If	
				Set oVBSFramework = New clsVBSFramework			
				oVBSFramework.Main(sLocalSpreadSheetPath)
			End If
										
		Else		    		   
		    Reporter.ReportEvent 1, "File not found", "File in path: " & sVBSFrameworkDir & "\clsVBSFramework.vbs " & " does not exist"
		End If
				
	Case VAPI_XP_TEST
	
		If oFS.FileExists(sFilePath) Then
			Set oFile = oFS.OpenTextFile(sFilePath, 1, False)
			sScript = oFile.ReadAll
			oFile.Close
			Set oFile = Nothing
			ExecuteGlobal sScript 'load vbs to global scope
			Set oVBSFramework = New clsVBSFramework
			oVBSFramework.Main oVBSFramework.GetTestSpreadsheetFromQC()
			'oVBSFramework.Main(sLocalSpreadSheetPath)
		Else
       		TDHelper.AddStepToRun "File not found", "File in path: " & sFilePath & "doesn not exists",,,"Failed"
		End If
		
	Case CMD_TEST
	
		If oFS.FileExists(sFilePath) Then
			Set oFile = oFS.OpenTextFile(sFilePath, 1, False)
			sScript = oFile.ReadAll
			oFile.Close
			Set oFile = Nothing
			ExecuteGlobal sScript 'load vbs to global scope
			Set oVBSFramework = New clsVBSFramework	
			
			'its posibble to run from comand line by add argument path to spreadsheet
			Set arrArgs = WScript.Arguments
			
			If arrArgs.Count = 1 Then ' run only one sheet with defaults
				sLocalSpreadSheetPath = oFS.GetAbsolutePathName(CStr(arrArgs.Item(0)))
				oVBSFramework.Main(sLocalSpreadSheetPath)
			Else
				If arrArgs.Count = 3 Then ' run sheet with another settings
					
					sLocalSpreadSheetPath = oFS.GetAbsolutePathName(CStr(arrArgs.Item(0)))
					globResultPath =  oFS.GetAbsolutePathName(CStr(arrArgs.Item(1)))
					globDataPath = oFS.GetAbsolutePathName(CStr(arrArgs.Item(2)))
					oVBSFramework.Main(sLocalSpreadSheetPath)
					
				Else
					'single run				
					oVBSFramework.Main(sLocalSpreadSheetPath)
				End If
					
			End If			
					
		Else       	  			
    		AppendToFile "c:\VBSFramework\Temp\" & "Automation.Log", Now() & " " & "File not found" & " Description: " & "File in path: " & sVBSFrameworkDir & "\clsVBSFramework.vbs " & "does not exists" & " Status: " & "Failed" & vbLf
		End If	
		
	Case Else 
		MsgBox "Not proper test type asigned"	
		
End Select

Set oFS = Nothing

Set oVBSFramework = Nothing
Set globCurrentTSTest =  Nothing
Set globCurrentTestSet = Nothing
Set globCurrentRun  = Nothing
Set globTDConnection  = Nothing	

SetLocale iOriginalLocale



'return value of run - form cmd-runner reason
If globRunMode = CMD_TEST Then
	Dim oStdOut
	Set oStdOut = WScript.StdOut
	oStdOut.WriteLine "exectime@" & globTestStart 
	oStdOut.WriteLine "duration@" & iDuration
	oStdOut.Close
	Set oStdOut = Nothing		
	Wscript.Quit globRetVal
ElseIf globRunMode = QTP_LOCAL_TEST Then	
	TestArgs("ExecTime") = globTestStart
	TestArgs("Duration") = iDuration	
End If

