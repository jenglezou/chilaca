Option Explicit

'==================================================================================================
Dim sLocalSpreadSheetPath : sLocalSpreadSheetPath = "C:\Chilaca\tests\Examples.xls"

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
sVBSFrameworkDir = "c:\chilaca"

Dim globRunMode ' global variable indicates runMode
Dim globRetVal ' return value of framework run
Dim globResultPath '  path where result spreadsheet will be stored
Dim globDataPath '  path where are data spreadsheets stored
Dim globTestStart 'time when spreadsheet started execution
Dim globDebug: globDebug = False ' if debug mode is on/of
Dim iDuration '  duration of test execution

'Constants
Const QTP_TEST = 1
Const VAPI_XP_TEST = 2
Const CMD_TEST = 3 
Const QTP_LOCAL_TEST = 4

Set oFS = CreateObject("Scripting.FileSystemObject")

'set to default 
globRunMode = -1

iOriginalLocale = SetLocale("en-gb")  'Get the test script location

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

sFilePath = oFS.GetAbsolutePathName(sVBSFrameworkDir & "\clsVBSFramework.vbs")

ExecuteFile sFilePath  'load vbs to global scope
Set oVBSFramework = New clsVBSFramework			

'************************************************************************************************************************************************
'**************** ADD THE REGISTRATION CODE FOR TESTOBJECT BELOW **********************************
'Registration code 
'Dim oExamplesLocal				'Don't need this as declared below in testobject section 									
Set oExamples = New clsExamples
oVBSFramework.oTestObjects.Add "EXAMPLES", oExamples
'************************************************************************************************************************************************

'Run
oVBSFramework.Main(sLocalSpreadSheetPath)

Set oFS = Nothing

'Set oVBSFramework = Nothing
'Set globCurrentTSTest =  Nothing
'Set globCurrentTestSet = Nothing
'Set globCurrentRun  = Nothing
'Set globTDConnection  = Nothing	

SetLocale iOriginalLocale


'************************************************************************************************************************************************
'**************** ADD THE TESTOBJECT CODE BELOW ***************************************************
'************************************************************************************************************************************************

'==================================================================================================
' clsExamples.vbs 
'
' Purpose: 
'==================================================================================================
'
'--------------------------------------------------------------------------------------------------
' AMENDEMENT HISTORY
'--------------------------------------------------------------------------------------------------
' Reason:Initial Version       Author:         Date:
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
' DESCRIPTION
'--------------------------------------------------------------------------------------------------
'**
' Describe change detail, include name of function amended
'**

'--------------------------------------------------------------------------------------------------
' Constants
'--------------------------------------------------------------------------------------------------

'--------------------------------------------------------------------------------------------------
' Class Start
'--------------------------------------------------------------------------------------------------
Class clsExamples							
	'==============================================================================================
	' CLASS VARIABLES
	'==============================================================================================

	'----------------------------------------------------------------------------------------------
	' Public Variables
	'----------------------------------------------------------------------------------------------

	'----------------------------------------------------------------------------------------------
	' Private Variables
	'----------------------------------------------------------------------------------------------
	Private sDemoMode 
	
	'----------------------------------------------------------------------------------------------
	' CLASS_INITIALIZE
	'----------------------------------------------------------------------------------------------
	Private Sub Class_Initialize
		sDemoMode = "OFF"
	End Sub

	'==============================================================================================
	' CLASS SUBS & FUNCTIONS
	'==============================================================================================
	'==============================================================================================
	' Function/Sub: RowDispatch
	' Purpose:		RowDispatch (must be a public function) is what the VBS framework uses to 
	'				access the Test Object's actions. 
	'
	' Parameters:	A spreadsheet row object. 
	'
	' Returns:		A value from the following list:
	'				XL_DISPATCH_PASS, XL_DISPATCH_FAIL, XL_DISPATCH_FAILCONTINUE,
	'				XL_DISPATCH_UNKNOWN, XL_DISPATCH_END, XL_DISPATCH_SKIP, XL_DISPATCH_CANCEL       
	'==============================================================================================
	Public Function RowDispatch(oRow, sLog, sOutputParams)
        Dim iRetVal, sKeyword

		Const sRoutine = "clsExamples.RowDispatch"
			
		oVBSFramework.oTraceLog.Entered(sRoutine & ", Keyword=""" & CStr(oRow.Cells(1, XL_KEYWORD).Value) & """")
		
        'Get keyword from spreadsheet row as a string
        sKeyword= CStr(oRow.Cells(1, XL_KEYWORD).Value)
        'Get the action string after the dot
        sKeyword = right(sKeyword, len(sKeyword)-inStrRev(sKeyword, "."))		
		
		Select Case uCase(sKeyword)
			'Dispatch known keyword functions		
			Case "SETOUTPUTPARAMS"							
				iRetVal = SetOutputParams(oRow,sOutputParams,sLog)
				iRetVal = ListCurrentRow(oRow, sLog, sOutputParams, XL_DISPATCH_PASS)
			Case "SHOWOUTPUTPARAMS"							
				'iRetVal = SetOutputParams(oRow,sOutputParams,sLog)
				iRetVal = ListCurrentRow(oRow, sLog, sOutputParams, XL_DISPATCH_PASS)
			Case "SAYHELLO"							
				iRetVal = SayHello(oRow,sOutputParams,sLog)
				iRetVal = ListCurrentRow(oRow, sLog, sOutputParams, XL_DISPATCH_PASS)
			Case Else
				iRetVal = XL_DISPATCH_UNKNOWN 	' Unknown keyword
		End Select

		RowDispatch = iRetVal
		oVBSFramework.oTraceLog.Exited(sRoutine & ", Keyword=""" & CStr(oRow.Cells(1, XL_KEYWORD).Value) & """")
	End Function

	'==============================================================================================
	' Function/Sub: SetOutputParams(oRow,sOutputParams,sLog)
	'==============================================================================================
	Private Function SetOutputParams(oRow,sOutputParams,sLog)					
		Dim sData,iRetVal
		Const sRoutine = "clsExamples.SetOutputParams" 
		oVBSFramework.oTraceLog.Entered(sRoutine)
		
		sData = CStr(oRow.Cells(1, XL_PARM_001).Value)		
		
		iRetVal = XL_DISPATCH_PASS
		
		sOutputParams  = sData
		SetOutputParams = iRetVal	
						
		oVBSFramework.oTraceLog.Exited(sRoutine)
	End Function
	'==============================================================================================
	' Function/Sub: HelloWorld(oRow,sOutputParams,sLog)
	'==============================================================================================
	Private Function SayHello(oRow,sOutputParams,sLog)					
		Dim sName,iRetVal
		Const sRoutine = "clsExamples.SayHello" 
		oVBSFramework.oTraceLog.Entered(sRoutine)
		
		iRetVal = XL_DISPATCH_PASS
		sName = CStr(oRow.Cells(1, XL_PARM_001).Value)		
		
		If sName = "" Then
			MsgBox "Hello World!"
		Else
			MsgBox "Hello " & sName & "!"
		End If
		
		SayHello = iRetVal							
		oVBSFramework.oTraceLog.Exited(sRoutine)
	End Function
	
	'==============================================================================================
	' Function/Sub:
	'==============================================================================================
	Private Function ListCurrentRow(oRow, sLog, sOutputParams, iRetVal)
		Dim sParam, sKeyword
		
		Const sRoutine = "clsExamples.ListCurrentRow"
		oVBSFramework.oTraceLog.Entered(sRoutine)
		
		sKeyword = CStr(oRow.Cells(1, XL_KEYWORD).Value)
		if iRetVal <> XL_DISPATCH_PASS Then
			iRetVal = XL_DISPATCH_FAILCONTINUE
			sLog = sLog & "E: Keyword " & sKeyword & " is pending development." & vbLf
		End If
		
		Dim i, sRowInput
		sRowInput = sKeyword & vbLF
		for i = XL_PARM_001 to XL_PARM_010
			sParam = CStr(oRow.Cells(1, i).Value)
			if trim(sParam) = "" then exit for
			sRowInput = sRowInput & "Input " & i - XL_PARM_001 + 1 & ":" & sParam & vbNewLine
		next
			
		If sDemoMode = "ON" Then 
			PopUp sRowInput, 1, "Row:" & oRow.Cells(1, i).row
			'msgbox sRowInput
		End if
		
		sLog = sLog & replace(sRowInput, vbNewline, ", ") & vbLF

		oVBSFramework.oTraceLog.Exited(sRoutine)
		ListCurrentRow = iRetVal
	End Function
	

'--------------------------------------------------------------------------------------------------
' Class End clsExamples
'--------------------------------------------------------------------------------------------------
End Class

'Registration code 
Public oExamples										
if not oVBSFramework.oTestObjects.isLoaded("EXAMPLES") then
	Set oExamples = New clsExamples		
	oVBSFramework.oTestObjects.Add "EXAMPLES", oExamples
End If

