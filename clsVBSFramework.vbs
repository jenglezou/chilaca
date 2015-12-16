'=========================================================================================================
'==================================================================================================
' clsVBSFramework.vbs
' Contains class clsVBSFramework,..
'--------------------------------------------------------------------------------------------------
' AMENDEMENT HISTORY
'--------------------------------------------------------------------------------------------------
'--------------------------------------------------------------------------------------------------

'--------------------------------------------------------------------------------------------------
' DESCRIPTION
'--------------------------------------------------------------------------------------------------
'**
' Describe change detail, include name of function amended
'**
'--------------------------------------------------------------------------------------------------
'==================================================================================================

' Forces explicit declaration (DIM) of all variables
' An error will occur if variables are used before being declared

Option Explicit

'==================================================================================================
' Constants
'==================================================================================================

' Paths
Dim PATH_RESOURCES : PATH_RESOURCES = sVBSFrameworkDir & "\temp\"
Dim PATH_TESTS : PATH_TESTS =  sVBSFrameworkDir & "\tests\"
Dim PATH_HOSTLOGFILE : PATH_HOSTLOGFILE = sVBSFrameworkDir & "\trace.log"
 
' Return codes for dispatcher and test functions
Const XL_DISPATCH_UNKNOWN = 1
Const XL_DISPATCH_END     = 2
Const XL_DISPATCH_SKIP    = 3
Const XL_DISPATCH_PASS    = 4
Const XL_DISPATCH_FAIL    = 5
Const XL_DISPATCH_CANCEL  = 6
Const XL_DISPATCH_FAILCONTINUE  = 7

Const ENV_CONFIG_NAME = "env_configuration"

Const LOG_ERROR = "ERROR"
Const LOG_WARNING = "WARNING"
Const LOG_SLOG = "LOG"
Const LOG_MESSAGE = "" 'new but supports back comp.
Const LOG_DEBUG = "DEBUG" '??? consider it

'==================================================================================================
' Global variables
'==================================================================================================

'spreadsheet column positions
Dim XL_DISABLE
Dim XL_RESULT
Dim XL_STREAM
Dim XL_KEYWORD
Dim XL_COMMENT
Dim XL_LOG
Dim XL_REFERENCE
Dim XL_OUTPUT_PARAMS
Dim XL_PARM_001, XL_PARM_002, XL_PARM_003, XL_PARM_004
Dim XL_PARM_005, XL_PARM_006, XL_PARM_007, XL_PARM_008
Dim XL_PARM_009, XL_PARM_010

Dim oQRSDataObject
Dim sVBSFrameworkDir
Dim oQC						' QC object
Dim dFrameworkStart			' Time when the test starts

'logging variables
Dim gTraceLogDepth
Dim gTraceLogOn
Dim gTraceLogInstanceNumber

gTraceLogDepth = 0
gTraceLogOn = True
gTraceLogInstanceNumber = 0

Dim sXUnitFileText, sXUnitTestSuiteName, sXUnitTestClassName, sXUnitStepRow, sXUnitStepName
'==================================================================================================
' Class Start
'==================================================================================================

Class clsVBSFramework

	'==============================================================================================
	' CLASS VARIABLES
	'==============================================================================================

	'----------------------------------------------------------------------------------------------
	' Public Variables
	'----------------------------------------------------------------------------------------------

	'----------------------------------------------------------------------------------------------
	' Private Variables
	'----------------------------------------------------------------------------------------------

	Public oXlApp					' Excel Application object
	Private oFso					' FileSystem object for file system operations
	'Private oQTP					' QTP Automation Qbject Model (Used to load .tsr file)
	
	Public oTestObjects				' Object to hold application objects
	Public oTraceLog				' Provides tracing for debug purposes
	Public oModules
	
	Public bPerformance
	Public sPerformanceFile			' Performance file
	Public sPerformanceAddInfo		' Performance additional info
	Public sDataSetPath				' HPQC/ALM Execution Path or defined DataSet folder/path
	Public bDataSetPrompt			' Flag to determine if prompt for DataSet value
	
	Private bLogLevel				' Flag to enable / disable desktop alerts
	Private sLogFileName			' Variable to hold log file name
	Private bAnimate				' Flag to enable spreadsheet animation
	Private bAutoCleanUp			' Flag to enable automatic termination of processes open during test execution
	Private bQCAutoDefect			' Flag to enable automatic generation of defects in QC 
	Private sQCAssignee				' Username(s) for QC defects and mails
	Private sQCEmailRecipients		' Username(s) for QC defect mails
	
	Private dictStreams				' Dictionary object to handle multiple streams in one spreadsheet

	Private dictExternalData		' Pre-cached content of external spreadsheets in form of Dictionary object (to speed-up resolving)
	
	Private sExcelProcesID
	
	'==============================================================================================
	' CLASS PROPERTIES
	'==============================================================================================

	'==============================================================================================
	' CLASS SUBS & FUNCTIONS
	'==============================================================================================

	'----------------------------------------------------------------------------------------------
	' CLASS_INITIALIZE
	'----------------------------------------------------------------------------------------------
	Private Sub Class_Initialize
		' Initialize random number generator
		Randomize Timer
		
		' Set Current Test Start Date and Time
		dFrameworkStart = Now()
		
		
		Set oFso = CreateObject("Scripting.FileSystemObject")
		
		Set dictStreams = CreateObject("scripting.dictionary")
		Set dictExternalData = CreateObject("scripting.dictionary")

		Set oTestObjects = New clsTestObjects
		
		Set oQC = New clsQC
		
		Set oTraceLog = New clsTraceLog		
		
		
		oTraceLog.HostLogMessage "START"
			
		
		Set oModules = New clsModules
			
		
		sLogFileName = "VBSFWExecutionLog.txt"

		bAnimate = False
		bAutoCleanUp = True 'False 
		bQCAutoDefect = False
		bPerformance = False
		bDataSetPrompt = True' REM what it is?
		sQCAssignee	= ""
		sQCEmailRecipients = ""
		
		bLogLevel = False
	
		On Error Resume Next
		If IsObject(oQRSData) then 
			Set oQRSDataObject = oQRSData
		End If
		
		If Err.Number <> 0 then
			Set oQRSDataObject = New clsDummyQRSData
		End If
		
		oQRSDataObject.dQRSFrameworkStart = dFrameworkStart
		
		On Error Goto 0
	End Sub

	'----------------------------------------------------------------------------------------------
	' CLASS_TERMINATE
	'----------------------------------------------------------------------------------------------
	Private Sub Class_Terminate
		' Remove references to objects
		Set oFso = Nothing
'		Set oXlApp = Nothing
'		If Not oQTP Is Nothing Then
'			oQTP.Quit
'			Set oQTP = Nothing	
'		End If	
		'Set oQTP = Nothing		
		Set oTestObjects = Nothing
		Set oTraceLog = Nothing	
		Set oModules = Nothing
		Set dictStreams = Nothing
		Set dictExternalData = Nothing
		Set oQC = Nothing
	End Sub

	'==============================================================================================
	' Function:		RowDispatch
	' Purpose:		Determines record_type and calls appropriate dispatcher
	'
	' Parameters:	oRow (current active row)
	'
	' Returns:		XL_DISPATCH_SKIP or XL_DISPATCH_UNKNOWN
	'
	'==============================================================================================
	Private Function RowDispatch(oRow)
		Dim iRetVal, lTimerStart, lTimerEnd, iCellColumnOffset
		Dim sKeyword, arrInputParams(10), sLog, sOutputParams, sScreenshotFullPath,sInputParams
		Dim oTestObject
		Dim arrStreams, sStream, sKey, sTimeStamp
		Dim i
		Dim tempsLog: tempsLog = ""
		
		Const sRoutine = "clsVBSFramework.RowDispatch"
		
		sInputParams = CStr(oRow.Cells(1, XL_PARM_001)) & "," & _
					   CStr(oRow.Cells(1, XL_PARM_002)) & "," & _
					   CStr(oRow.Cells(1, XL_PARM_003)) & "," & _
					   CStr(oRow.Cells(1, XL_PARM_004)) & "," & _
					   CStr(oRow.Cells(1, XL_PARM_005)) & "," & _
					   CStr(oRow.Cells(1, XL_PARM_006)) & "," & _
					   CStr(oRow.Cells(1, XL_PARM_007)) & "," & _
					   CStr(oRow.Cells(1, XL_PARM_008)) & "," & _
					   CStr(oRow.Cells(1, XL_PARM_009)) & "," & _
					   CStr(oRow.Cells(1, XL_PARM_010))
								 
		
		oTraceLog.Entered(sRoutine & " Keyword=" & CStr(oRow.Cells(1, XL_KEYWORD)) & ", Params=" & sInputParams)
				
		sLog = "" 
		oQRSDataObject.sQRSLog = ""
		sOutputParams = "" 
		oQRSDataObject.sQRSOutputParams = ""

		'skip row processing if all rowstreams are failed
		' TODO traceloging of streams
		If dictStreams.Count > 0 Then
			arrStreams = Split(oRow.Cells(1, XL_STREAM).Text, ",")
			For Each sStream In arrStreams
				If dictStreams(sStream) = "F" Then
					WriteIntoLogAndOutputParamsCells oRow, "I: Row skipped because stream [" & sStream & "] has already failed." & vblf, ""
					RowDispatch = XL_DISPATCH_SKIP
					oTraceLog.Exited(sRoutine & ", Keyword=" & CStr(oRow.Cells(1, XL_KEYWORD)))
					Exit Function
				ElseIf dictStreams(sStream) = "S" Then
					WriteIntoLogAndOutputParamsCells oRow, "I: Row skipped because stream [" & sStream & "] has skip flag." & vblf, ""
					RowDispatch = XL_DISPATCH_SKIP
					oTraceLog.Exited(sRoutine & ", Keyword=" & CStr(oRow.Cells(1, XL_KEYWORD)))
					Exit Function
				End If
			Next
		End If

		' Check if row disabled
		'If (oRow.Cells(1, XL_DISABLE).Value) or (not ltrim(oRow.Cells(1, XL_RESULT).Value) = "") Then
		If oRow.Cells(1, XL_DISABLE).Value <> "" Then
			iRetVal = XL_DISPATCH_SKIP
			oTraceLog.Message "SKIPPED", LOG_MESSAGE
		Else
						
			If DataAndConfigParamsRowProcessing(oRow, sLog) = False Then
				iRetVal = XL_DISPATCH_FAIL				
			Else
				sKeyword = CStr(oRow.Cells(1, XL_KEYWORD).Value)
				If InStr(sKeyword, ".") > 0 Then
					If globDebug Then
						oTraceLog.Message "Calling test object key word", LOG_DEBUG
					End if
					'loading module
					
					oModules.Load(UCase(mid(sKeyword, 1, InStr(sKeyword, ".") - 1)))
					
					' get test object, if its alias we get name of test object	
							
					Set oTestObject = oVBSFramework.oTestObjects.Item(oModules.GetNameOfTestObject(UCase(mid(sKeyword, 1, InStr(sKeyword, ".") - 1))))
					
					if oTestObject is Nothing Then				
						oTraceLog.Message "Application not found: " & mid(sKeyword, 1, InStr(sKeyword, ".") - 1), LOG_ERROR
						sLog = sLog & " E: Application not found: " & mid(sKeyword, 1, InStr(sKeyword, ".") - 1)
						iRetVal = XL_DISPATCH_FAIL
					Else
						lTimerStart = Timer
						
						'Execute the action
'						
						oTraceLog.Message "XLS: " & oRow.Parent.parent.name , LOG_MESSAGE
						'oTraceLog.Message " Keyword: " & sKeyword, LOG_MESSAGE
																						
						iRetVal = oTestObject.RowDispatch(oRow, sLog, sOutputParams)	'Can't use QRS obj.variables here because they get lost on return
						lTimerEnd = Timer
						
						If iRetVal = XL_DISPATCH_UNKNOWN Then
							sLog = sLog & "E: Unknown keyword" & vblf
						End If
						
						sLog = sLog & "D: Execution time: " & Round(lTimerEnd - lTimerStart) & " s" & vblf
						oQRSDataObject.sQRSLog = sLog						
						
					'	WriteIntoLogAndOutputParamsCells oRow, sLog, sOutputParams						
					'	oTraceLog.Message sLog, LOG_SLOG 					
																	
						oQRSDataObject.sQRSLog = sLog 									'For recovery
						oQRSDataObject.sQRSOutputParams	= sOutputParams					'For recovery
					
					'	ProcessOutputParams oRow
						
						Set oTestObject = Nothing
					End if
				Else
					'iRetVal = XL_DISPATCH_UNKNOWN
					' Get keyword from spreadsheet as string
					'sKeyword = CStr(oRow.Cells(1, XL_KEYWORD).Value)
					If oRow.Cells(1, XL_KEYWORD).Value <> "" Then
						If globDebug Then
							oTraceLog.Message "Calling framework key word", LOG_DEBUG
						End if
					'	oTraceLog.Message "> Keyword: " & sKeyword, LOG_MESSAGE	
						Select Case uCase(sKeyword)
		
							' Dispatch known config functions
							'Case "<keyword>"			iRetVal = K_<Keyword>(oRow)
							Case "PAUSE", "CONTINUE", "PAUSESHEET"
								Select Case Msgbox("Press Yes to continue with this sheet, No to end this Sheet or Cancel to end the Test.", _
													vbYesNoCancel, "Continue Sheet - " & oRow.Parent.Name)
									Case vbYes		iRetVal = XL_DISPATCH_SKIP
									Case vbNo		iRetVal = XL_DISPATCH_END
									Case vbCancel	iRetVal = XL_DISPATCH_CANCEL
								End Select
							Case "PAUSEROW"
								iRetVal = XL_DISPATCH_SKIP
								Select Case Msgbox("Press Yes to run the next row, No to skip the next step or Cancel to end the Test.", _
													vbYesNoCancel, "Run next row - " & oRow.Cells(2, XL_KEYWORD).Value)
									Case vbCancel	iRetVal = XL_DISPATCH_CANCEL
									Case vbNo		oRow.Cells(2, XL_DISABLE).Value = "TRUE"
								End Select
							Case "CONTINUESHEETIF"
								iRetVal = XL_DISPATCH_PASS
								if not UCase(CStr(oRow.Cells(1, XL_PARM_001).Value)) = "TRUE" then
									iRetVal = XL_DISPATCH_END
								end if
							Case "SETOUTPUT"
								Select Case Msgbox("Press Yes for " & CStr(oRow.Cells(1, XL_PARM_001).Value) & _
													", No for " & CStr(oRow.Cells(1, XL_PARM_002).Value) & " or Cancel to end the Test.", _
													vbYesNoCancel, "Set Output")
									Case vbYes
										oRow.Cells(1, XL_LOG).Value = oRow.Cells(1, XL_PARM_001).Value
										iRetVal = XL_DISPATCH_PASS
									Case vbNo
										oRow.Cells(1, XL_LOG).Value = oRow.Cells(1, XL_PARM_002).Value
										iRetVal = XL_DISPATCH_PASS
									Case vbCancel	iRetVal = XL_DISPATCH_CANCEL
								End Select
							Case "SELECTITEM"
								iRetVal = K_SelectItem(oRow)
							
							'TODO: extend functionality (multiple params?, error if output_params cell hasn't got name defined, ..)
							Case "SETPARAMETER", "PARAMETER"
							
								'process first param (to allow empty string as param value)
								sOutputParams = "PARAM1=" & Replace(Replace(oRow.Cells(1, XL_PARM_001).Value, ",", "^"), "=", "~") & ","
								'oRow.Cells(1, XL_OUTPUT_PARAMS).Value = "PARAM1=" & Replace(Replace(oRow.Cells(1, XL_PARM_001).Value, ",", "^"), "=", "~") & ","
							
								'WriteIntoLogAndOutputParamsCells oRow, sLog, sOutputParams
								
								'process other params
								iCellColumnOffset = 1
								Do Until oRow.Cells(1, XL_PARM_001 + iCellColumnOffset).Value = ""
								'	oRow.Cells(1, XL_OUTPUT_PARAMS).Value = oRow.Cells(1, XL_OUTPUT_PARAMS).Value & "PARAM" & (iCellColumnOffset + 1) & "=" & Replace(Replace(oRow.Cells(1, iCellColumnOffset + XL_PARM_001).Value, ",", "^"), "=", "~") & ","
									sOutputParams = sOutputParams & "PARAM" & (iCellColumnOffset + 1) & "=" & Replace(Replace(oRow.Cells(1, iCellColumnOffset + XL_PARM_001).Value, ",", "^"), "=", "~") & ","
									iCellColumnOffset = iCellColumnOffset + 1
								Loop
								
							'	ProcessOutputParams oRow
												
								iRetVal = XL_DISPATCH_PASS
								
							Case "STREAMS", "STREAMSTORUN"
								If dictStreams.Count > 0 Then
									arrStreams = Split(oRow.Cells(1, XL_PARM_001).Text, ",")
									
									'skip if defined StreamToRun doesn't exist in spreadsheet
									For Each sStream In arrStreams
										If Not dictStreams.Exists(sStream) Then
											iRetVal = XL_DISPATCH_SKIP 'TODO: should be FAIL?
											sLog = sLog & "E: Stream [" & sStream & "] not found in spreadsheet." & vblf
											WriteIntoLogAndOutputParamsCells oRow, sLog, ""
											oTraceLog.Exited(sRoutine & ", Keyword=" & CStr(oRow.Cells(1, XL_KEYWORD)))
											Exit Function
										End If
									Next
									
									'mark streams not defined in StreamsToRun with skip flag
									'process all streams if nothing is defined
									If UBound(arrStreams) >= 0 Then
										For Each sKey In dictStreams.Keys
											If Not ArrayContainsItem(arrStreams, sKey) Then
												dictStreams(sKey) = "S"
												sLog = sLog & "I: Stream [" & sKey & "] skip flag set." & vblf
											End If
										Next
									End If
									
									WriteIntoLogAndOutputParamsCells oRow, sLog, ""
									iRetVal = XL_DISPATCH_PASS
								Else
									'skip if no streams defined in spreadsheet
									sLog = sLog & "No streams defined in spreadsheet. Skipping." & vblf
									iRetVal = XL_DISPATCH_SKIP
								End If
	' REM changedto ScreenShot - it is in function.vbs
							Case "CAPTURESCREEN", "CAPTURESCREENSHOT", "SCREENCAPTURE", "SCREENSHOT"
Rem								 sScreenshotFullPath = globResultPath & "\" & "screen" & YYYYMMDDHHMMSS(Now()) & ".png"
								 'Desktop.CaptureBitmap sScreenshotFullPath								 
								 Screenshoot sScreenshotFullPath, true
								 
								 oTraceLog.Message "Screen captured. Screenshoot full path:" & sScreenshotFullPath, LOG_MESSAGE
								 
								'WriteIntoLogAndOutputParamsCells oRow, "I: Screen captured: " & sScreenshotFullPath, ""								 	
	
								 iRetVal = XL_DISPATCH_PASS						
							Case "ANIMATE"
								bAnimate = True
								oXlApp.Visible = True
								oRow.Parent.Activate							
								sLog = sLog & "I: Test is running by VBSFramework" & vblf
							'	WriteIntoLogAndOutputParamsCells oRow, sLog, ""
								iRetVal = XL_DISPATCH_PASS
								
							Case "AUTOCLEANUP", "CLEANUP"
								Select Case UCase(CStr(oRow.Cells(1, XL_PARM_001).Value))
									Case "OFF", "FALSE", "NO"
										bAutoCleanUp = False
									Case Else
										bAutoCleanUp = True
								End Select
								oQRSDataObject.bQRSAutoCleanUp = bAutoCleanUp				'For recovery
								iRetVal = XL_DISPATCH_PASS
								
							Case "QCAUTODEFECT"
								Select Case UCase(CStr(oRow.Cells(1, XL_PARM_001).Value))
								Case "OFF", "FALSE", "NO"
									bQCAutoDefect = False						
								Case Else
									bQCAutoDefect = True
									sQCEmailRecipients = UCase(CStr(oRow.Cells(1, XL_PARM_002).Text))
									sQCAssignee = UCase(CStr(oRow.Cells(1, XL_PARM_003).Text))
								End Select
								oQRSDataObject.bQRSQCAutoDefect = bQCAutoDefect				'For recovery
								oQRSDataObject.sQRSQCEmailRecipients = sQCEmailRecipients	'For recovery
								oQRSDataObject.sQRSQCAssignee = sQCAssignee					'For recovery
								iRetVal = XL_DISPATCH_PASS
		
							Case "SELECTDATATABLE", "SELECTXLS"
								iRetVal = k_SelectDataTable(oRow)
							Case "SYSTEMCOMMAND"
								'Haven't tested this but it WILL be useful
								'WshShell.Run CStr(oRow.Cells(1, XL_PARM_001).Value), 1, False
								iRetVal = XL_DISPATCH_PASS
							Case "LOGLEVEL"			iRetVal = K_LogLevel(oRow)
							Case "LOADAPP","LOADVBS", "TESTOBJECT"	iRetVal = LoadVBS(oRow)
							Case "LOADMODULE","MODULE", "LOAD"	iRetVal = LoadModuleAction(oRow)
							Case "DELAY"			iRetVal = K_Delay(oRow)
							Case "EXEC","LOADXLS","XLOPEN","OPENXLS"
								iRetVal = K_Exec(oRow)
							Case "SAVE"			iRetVal = K_Save(oRow)
							Case "END"			iRetVal = XL_DISPATCH_END
							Case "COMMENT", ""	iRetVal = XL_DISPATCH_SKIP
							Case "DESIGNMODEUPDATEATTACHMENTS"
								iRetVal = DesignModeUpdateAttachmentsAction(oRow, sLog)
							'	WriteIntoLogAndOutputParamsCells oRow, sLog, sOutputParams
							Case "DESIGNMODEEND"
								iRetVal = DesignModeEnd(oRow, sLog)
							'	WriteIntoLogAndOutputParamsCells oRow, sLog, sOutputParams
							Case "EXECUTESTATEMENT"
								iRetVal = ExecuteStatement(oRow, sLog)
							'	WriteIntoLogAndOutputParamsCells oRow, sLog, sOutputParams
							Case "RUNASUSER"
								SetRunAsUser(UCase(Trim(CStr(oRow.Cells(1, XL_PARM_001).Value))))
								sLog = sLog & "I: Tests will be executed as user " & GetRunAsUser()
								iRetVal = XL_DISPATCH_PASS
							'	WriteIntoLogAndOutputParamsCells oRow, sLog, sOutputParams
							Case "PERFORMANCE"							
								Select Case UCase(CStr(oRow.Cells(1, XL_PARM_001).Value))
								Case "OFF", "FALSE", "NO"
									bPerformance = False
									sPerformanceAddInfo = ""						
								Case Else
									bPerformance = True
									sPerformanceFile = CStr(oRow.Cells(1, XL_PARM_002).Value)
									sPerformanceAddInfo = CStr(oRow.Cells(1, XL_PARM_003).Value)
									If sPerformanceFile = "" Then
										sPerformanceFile = PATH_RESOURCES & globCurrentTSTest.Test.Name & ".csv"'
									ElseIf InStr(sPerformanceFile,"\") > 0 Then
									Else
										sPerformanceFile = PATH_RESOURCES & sPerformanceFile
									End If
									WritePerformanceFile sPerformanceFile,"task,start,end,duration,testset", 1
								End Select
								iRetVal = XL_DISPATCH_PASS
							'	WriteIntoLogAndOutputParamsCells oRow, sLog, sOutputParams
							Case "FAIL" 'for testing of framework purpose
								iRetVal = XL_DISPATCH_FAIL	
							Case "SENDMAIL"			
								iRetVal = SendMail(oRow, sLog)
							'	ProcessOutputParams oRow
							Case "FINDMAIL"			
								iRetVal = FindMail(oRow, sLog)
							'	ProcessOutputParams oRow 
							'recording in QTP for future implementation	
							Case "RECORDING"
								iRetVal	= XL_DISPATCH_UNKNOWN
							Case "DEBUG"
							' turn on debug mode in spreadsheet
								oTraceLog.TurnOnDebug							
								iRetVal	= XL_DISPATCH_PASS							
							Case Else
								' Unknown keyword
								iRetVal = XL_DISPATCH_UNKNOWN
						End Select
					Else	
						iRetVal = XL_DISPATCH_SKIP
						oTraceLog.Message "SKIPPED - Blank line", LOG_MESSAGE
					End If
				
					
				End If 'InStr(sKeyword, ".") > 0
			End If 'DataAndConfigParamsRowProcessing(oRow) = False
			
			'array for better step run report
			arrInputParams(0)=oRow.Cells(1, XL_PARM_001)
			arrInputParams(1)=oRow.Cells(1, XL_PARM_002)
			arrInputParams(2)=oRow.Cells(1, XL_PARM_003)
			arrInputParams(3)=oRow.Cells(1, XL_PARM_004)
			arrInputParams(4)=oRow.Cells(1, XL_PARM_005)
			arrInputParams(5)=oRow.Cells(1, XL_PARM_006)
			arrInputParams(6)=oRow.Cells(1, XL_PARM_007)
			arrInputParams(7)=oRow.Cells(1, XL_PARM_008)
			arrInputParams(8)=oRow.Cells(1, XL_PARM_009)
			arrInputParams(9)=oRow.Cells(1, XL_PARM_010)
			
			
			
			'oTraceLog.Message " Keyword: " & sKeyword, LOG_MESSAGE	
	
			' REM place where for me shoud be writing to xls sheet
			WriteIntoLogAndOutputParamsCells oRow, sLog, sOutputParams
			
			
			' proces output parameter with tempLog string
			ProcessOutputParams oRow, tempsLog
			
			WriteIntoLogAndOutputParamsCells oRow, tempsLog, " "
			
			If tempsLog <> "" Then
				sLog = sLog & tempsLog 
			End If 
		
			'write sLog to trace log file
			oTraceLog.Message sLog, LOG_SLOG
			' step message loaded to ALM
			Call oTraceLog.StepMessage(CStr(oRow.Cells(1, XL_KEYWORD)), iRetVal, sLog, arrInputParams, sOutputParams)

'***** For xUnit reporting ********
			sXUnitStepName = CStr(oRow.Cells(1, XL_KEYWORD))
			sXUnitStepRow = right("0" & oRow.Row, 2)
			if sXUnitStepName <> "" then
				sXUnitFileText = sXUnitFileText & "<testcase classname=""" & sXUnitTestClassName & """ name=""Step" & sXUnitStepRow & "-" & sXUnitStepName & """ time=""0"">" & vbNewLine 
				if iRetval = XL_DISPATCH_FAIL or iRetval = XL_DISPATCH_FAILCONTINUE then
					sXUnitFileText = sXUnitFileText & "<error type=""exception"" message=""error message"">" & vbNewLine & sLog & vbNewLine & "</error>" & vbNewLine
				end if 
				sXUnitFileText = sXUnitFileText & "</testcase>" & vbNewLine
			End If
'***** For xUnit reporting ********
		
		End If 'oRow.Cells(1, XL_DISABLE).Value <> ""
		
				
		RowDispatch = iRetVal

		oTraceLog.Exited(sRoutine & ", Keyword=" & CStr(oRow.Cells(1, XL_KEYWORD)))
	End Function

	'==============================================================================================
	' Subroutine:	Main
	' Purpose:		Execute the required tests by reading and processing each row of sWorkbook
	'
	' Parameters:	dierectory path and name of worksheet to execute
	'
	'==============================================================================================
   	Public Sub Main(sWorkbook)
		Dim oWorkbook, oSheet, oCurrentTest, oBug
		Dim sOutput, sTempMessage, sGridTitles, sGridValues, sParm, sExpectedResult, sKey, sDefectDescription
		Dim aInputData, aWorkbookName, arrStreams
		Dim i, iCounter, iRow, iResult
		Dim bFailure, bDefectFound, bDefectThisSheet, bContinue, bBoldInitial
		Dim lTimerStart,lTimerEnd 'for measuring duration of whole test
		Dim sDurationLog
		Dim sResultWorkBook
		Dim oShell
		Dim sUserName
			
		If sWorkbook = "" Then 			
			oTraceLog.HostLogMessage "Excel test spreadsheet not specified, or no Excel test spreadsheet attachment."	
			oTraceLog.HostLogMessage "END" 
			Exit Sub
		End If

		Set oXlApp = CreateObject("Excel.Application")
	
		'set position of excel application 
		oXlApp.WindowState = -4143'Normal (Allow resezing)
		oXlApp.Left = 20
		oXlApp.Top = 150
		oXlApp.Height = 400
		oXlApp.Width = 700
		'oXlApp.Visible = True 
		oXlApp.Visible = False 
		oXlApp.DisplayAlerts = False
		oXlApp.AskToUpdateLinks = False
		'oXlApp.ScreenUpdating = True

		On Error Resume Next
	 		oFso.createfolder (globResultPath)
	 		oFso.createfolder (globDataPath)
		On Error GoTo 0

		sXUnitTestSuiteName = "AutomatedTestSuite"
		sXUnitFileText = "<?xml version=""1.0"" encoding=""UTF-8""?>" & vbNewLine & _
						 "<testsuites>" & vbNewLine & _
						 "<testsuite name=""" & sXUnitTestSuiteName & """ tests=""[TESTS]"" errors=""[ERRORS]"" failures=""[FAILURES]"" skip=""[SKIP]"">" & vbNewLine
		
		If sWorkbook <> "" Then 
			'get general start of spreadsheet
			globTestStart = Now()
			
			If oFso.FileExists(sWorkbook) Then
				sXUnitTestClassName = sXUnitTestSuiteName & "." & oFso.GetBaseName(sWorkbook)
			
				If oFso.FolderExists(globResultPath) Then	
'msgbox globDataPath
				
					If oFso.FolderExists(globDataPath) Then	
						'get general foler result folder
						globResultPath = globResultPath & "\" & oFso.GetBaseName(sWorkbook) & "_RESULT" & YYYYMMDDHHMMSS(globTestStart)
						'get path of result workbook
						sResultWorkBook = GetResultWorkBookPath(sWorkbook, globResultPath & "\")
									
						'set trace log file path
						oTraceLog.LogFilePath = globResultPath & "\" & Mid(sResultWorkBook, InStrRev(sResultWorkBook, "\") + 1) & "_TRACELOG.txt"
						
						If globDebug Then
							oTraceLog.Message "Result work book: " & sResultWorkBook, LOG_DEBUG
							oTraceLog.Message "Trace log: " & globResultPath & "\" & Mid(sResultWorkBook, InStrRev(sResultWorkBook, "\") + 1) & "_TRACELOG.txt", LOG_DEBUG
						End If
					Else
						oTraceLog.Message "Defined Data folder " & globDataPath & " does not exists", LOG_ERROR
						oTraceLog.Message "END", LOG_MESSAGE
						oTraceLog.HostLogMessage "Defined Data folder " & globDataPath & " does not exists"	
						oTraceLog.HostLogMessage "END" 				
						oTraceLog.Message "Test Failed", LOG_ERROR
					
						If oQC.IsQCRun Then												
						'set test result to ALM
							globCurrentRun.Status = "Failed"						
							oTraceLog.StepMessage "ATTACHMENT PROBLEM", XL_DISPATCH_FAIL, "Defined Data folder " & globDataPath & " does not exists", "", "" 
						End if
					Exit Sub
						
					End If	
				Else
					oTraceLog.Message "Defined Result folder " & globResultPath & " does not exists", LOG_ERROR
					oTraceLog.Message "END", LOG_MESSAGE
					oTraceLog.HostLogMessage "Defined Result folder " & globResultPath & " does not exists"	
					oTraceLog.HostLogMessage "END" 				
					oTraceLog.Message "Test Failed", LOG_ERROR
					
					If oQC.IsQCRun Then												
						'set test result to ALM
						globCurrentRun.Status = "Failed"						
						oTraceLog.StepMessage "ATTACHMENT PROBLEM", XL_DISPATCH_FAIL, "Defined Result folder " & globResultPath & " does not exists", "", "" 
					End if
					Exit Sub
				End If	
				
			Else
				oTraceLog.Message "File " & sWorkbook & " does not exists", LOG_ERROR
				oTraceLog.Message "END" ,LOG_MESSAGE
				oTraceLog.HostLogMessage "Test has no correct excel attachment path"	
				oTraceLog.HostLogMessage "END" 				
				oTraceLog.Message "Test Failed", LOG_ERROR
				If oQC.IsQCRun Then												
					'set test result to ALM
					globCurrentRun.Status = "Failed"						
					oTraceLog.StepMessage "ATTACHMENT PROBLEM", XL_DISPATCH_FAIL, "File " & sWorkbook & " does not exists", "", "" 
				End if
				Exit Sub
			End If					
		End If	
		
		
		oTracelog.TurnOn()
		
		'add start of tracelog 

		oTraceLog.Message "START", LOG_MESSAGE 		
		'add additional info for run environment 
		Set oShell = CreateObject("WScript.Shell")
		sUserName = oShell.ExpandEnvironmentStrings("%USERNAME%")
		oTraceLog.Message "Logged user: " & sUserName, LOG_MESSAGE 
		Set oShell = Nothing
	

		
		bContinue = true
		'sDefectDescription = ""				'This string will hold text from all failed Rows.
		sDefectDescription = "Workbook:" & sWorkbook & vbnewline : oQRSDataObject.sQRSDefectDescription = sDefectDescription
		bDefectFound = False : oQRSDataObject.bQRSDefectFound = bDefectFound		'This will be true if a defect is to be created

		' Remove dynamically resolved data but keep values
		RemoveFormulasKeepValuesInWorkbook sWorkbook
		
		' Copy workbook to result directory
		'Workbook = CopyWorkbookOld(sWorkbook)								'Use QTP result folder
		 
		'sWorkbook = CopyWorkbook(sWorkbook, sVBSFrameworkDir & "\RESULTS\")	'Use the RESULTS folder in the current test folder
		sWorkbook = CopyWorkbook(sWorkbook, sResultWorkBook)	'Use the RESULTS folder in the current test folder
		
		oTraceLog.HostLogMessage "MAIN: " & "Test started for workbook " & sWorkbook
		oTraceLog.HostLogMessage "TRACE_LOG FILE: " & oTraceLog.LogFilePath
		
		oQRSDataObject.sQRSWorkbook = sWorkbook					'For recovery
		
		' Open workbook
		oTraceLog.Message "MAIN: " & "Test started for workbook " & sWorkbook, LOG_MESSAGE
		Set oWorkbook = OpenWorkbook(sWorkbook) 
		set oQRSDataObject.oQRSWorkbook = oWorkbook				'For recovery
	
		Dim oWMIService,oWMIDateTime,oProcessList,oProcess
		Set oWMIDateTime = CreateObject("WbemScripting.SWbemDateTime")
		oWMIDateTime.SetVarDate(dFrameworkStart)
		
		Set oWMIService = GetObject("winmgmts:{impersonationLevel=impersonate}!\\" & "." & "\root\cimv2")
		Set oProcessList = oWMIService.ExecQuery("Select * from Win32_Process WHERE CreationDate > '" & oWMIDateTime & "' AND name = 'excel.exe'")
		
		If oProcessList.Count = 1 Then
			For Each oProcess in oProcessList
				sExcelProcesID = oProcess.ProcessID 
			Next
		'	If globDebug Then
				oTraceLog.Message "Excel proces ID: " & sExcelProcesID, LOG_MESSAGE
		'	End If	
			'MsgBox sExcelProcesID
		Else
		'	MsgBox oProcessList.Count
		End If
		Set oWMIService = Nothing
		Set oWMIDateTime = Nothing
		Set oProcessList = Nothing
		
		If bAnimate Then oXlApp.Visible = True 'JE added

		
		If oQC.IsQCRun() Then
			'if test is run from QC and Active Field is set to "N" then do not execute
			oTraceLog.Message "ALM is running", LOG_MESSAGE
			Rem If globCurrentTSTest.Test.Field("TS_USER_TEMPLATE_01") = "N" Then
			
				Rem Set oCurrentTest = globCurrentTSTest.Test
				
				Rem oWorkbook.Worksheets.Item(1).Rows(2).Insert
				Rem oWorkbook.Worksheets.Item(1).Rows(2).Font.ColorIndex		= 2
	    		Rem oWorkbook.Worksheets.Item(1).Rows(2).Font.Size				= 11
				Rem oWorkbook.Worksheets.Item(1).Rows(2).Interior.ColorIndex	= 5
				Rem oWorkbook.Worksheets.Item(1).Rows(2).Interior.ColorIndex	= 3
				Rem oWorkbook.Worksheets.Item(1).Rows(2).Cells(1,1).WrapText 	= False
				Rem oWorkbook.Worksheets.Item(1).Rows(2).Cells(1,1).Value 		= "Unable to run - field Active set to 'N'."
				Rem oTraceLog.Message "Unable to run - field Active set to 'N'.", LOG_ERROR
				Rem bContinue = false
			Rem End If
		Else 
			'TODO add message according to run mode
			oTraceLog.Message "ALM is not running", LOG_MESSAGE	
		End If
		
		If bContinue Then
			'Read the ini file of source and dependencies
			oModules.SourceFilesAndDependenciesFromINI sVBSFrameworkDir & "\VBSFramework.ini", sOutput
			'load core modules
			oModules.LoadCoreModules

			If ValidateWorkbook(oWorkbook) Then
				' Loop through all sheets
				For Each oSheet in oWorkbook.Worksheets
					bDefectThisSheet = False
					If bAnimate then oSheet.Activate 'JE added
				
					'assign column numbers for currently processed sheet					
					AssignHeaderColumnXLValues oSheet
					
					'added fix for excel bug, when columnwidth is very small, setting rowheight works not correct, now is check only for Log row
					
					If oSheet.Columns(XL_LOG).ColumnWidth < 12 Then oSheet.Columns(XL_LOG).ColumnWidth  = 12 
					
					' Skip header row, start with row 2
					iRow = 2
					' Loop through all rows
					lTimerStart = Timer
					Do Until iRow > oSheet.UsedRange.Rows.Count
						'Save the initial state of bold for KEYWORD
						sExpectedResult = oSheet.Rows(iRow).Cells(1, XL_RESULT).Value 
						bBoldInitial = oSheet.Rows(iRow).cells(1, XL_KEYWORD).Font.Bold
	
						' Set result column for row in progress
						oSheet.Rows(iRow).Cells(1, XL_KEYWORD).Font.Bold     	  = True
						oSheet.Rows(iRow).Cells(1, XL_RESULT).Value               = "---->"
						oSheet.Rows(iRow).Cells(1, XL_RESULT).Font.ColorIndex     = -4105		' Automatic
						oSheet.Rows(iRow).Cells(1, XL_RESULT).Interior.ColorIndex = 45			' Orange
						oSheet.Rows(iRow).Cells(1, XL_RESULT).Interior.Pattern    = 1
	
						bFailure = False
					
						set oQRSDataObject.oQRSRow = oSheet.Rows(iRow)			'For recovery
						
						'Execute the action						
						iResult = RowDispatch(oSheet.Rows(iRow))
						
						set oQRSDataObject.oQRSRow = Nothing
						
						'make big rows smaller
						If oSheet.Rows(iRow).RowHeight > 75 Then oSheet.Rows(iRow).RowHeight = 75
						
						'Added 25/11/05
						'check to see if overriding of results is required
						'this option allows the user to override Return codes from the dispatcher and test functions
						sTempMessage = CStr(oSheet.Rows(iRow).Cells(1, XL_OUTPUT_PARAMS).Value)
						if InStr(sTempMessage, "XL_DISPATCH_") then
							'Override is required
							aInputData = Split(sTempMessage, vblf, -1)
							Select Case  aInputData(0)
								Case "XL_DISPATCH_UNKNOWN"		iResult = XL_DISPATCH_UNKNOWN
								Case "XL_DISPATCH_END"			iResult = XL_DISPATCH_END
								Case "XL_DISPATCH_SKIP"			iResult = XL_DISPATCH_SKIP
								Case "XL_DISPATCH_PASS"			iResult = XL_DISPATCH_PASS
								Case "XL_DISPATCH_FAIL"			iResult = XL_DISPATCH_FAIL
								Case "XL_DISPATCH_CANCEL"		iResult = XL_DISPATCH_CANCEL
								Case Else
							End Select
						End if
	
						'Reset to the initial state of bold for KEYWORD
						oSheet.Rows(iRow).cells(1, XL_KEYWORD).Font.Bold = bBoldInitial
						
						'for returning reason, we not return SKIP result
						If iResult <> XL_DISPATCH_SKIP Then
							If iResult <> XL_DISPATCH_END Then
								'set actuall global return value
								globRetVal = iResult
							End If
						End If	
						
						If iResult = XL_DISPATCH_UNKNOWN Then
							' Set failure flag
							bFailure = True
							'MsgBox "Unknown keyword: " & oSheet.name & "." & oSheet.Rows(iRow).Cells(1, XL_KEYWORD).Value & vbNewline & "Test will now exit."
							oTraceLog.Message "Unknown keyword: " & oSheet.name & "." & oSheet.Rows(iRow).Cells(1, XL_KEYWORD).Value, LOG_WARNING
						
							'oTraceLog.Message("Unknown keyword: " & oSheet.Rows(iRow).Cells(1, XL_KEYWORD).Value)
							If bQCAutoDefect then 
								if bDefectThisSheet = False then 
									sDefectDescription = sDefectDescription & "Worksheet:" & oSheet.name & vbnewline  : oQRSDataObject.sQRSDefectDescription = sDefectDescription
									bDefectThisSheet = True
								end if 
								sDefectDescription = sDefectDescription & "Row:" & iRow & vbnewLine & RowToText(oSheet.Rows(iRow))  : oQRSDataObject.sQRSDefectDescription = sDefectDescription
								bDefectFound = True : oQRSDataObject.bQRSDefectFound = bDefectFound
							End If
						
							' Set result column
							oSheet.Rows(iRow).Cells(1, XL_RESULT).Value               = "unknown"
							oSheet.Rows(iRow).Cells(1, XL_RESULT).Font.ColorIndex     = -4105		' Automatic
							oSheet.Rows(iRow).Cells(1, XL_RESULT).Interior.ColorIndex = 6			' Yellow
							oSheet.Rows(iRow).Cells(1, XL_RESULT).Interior.Pattern    = 1
						ElseIf iResult = XL_DISPATCH_END Then
							' Set result column
							oSheet.Rows(iRow).Cells(1, XL_RESULT).Value               = "end"
							oSheet.Rows(iRow).Cells(1, XL_RESULT).Font.ColorIndex     = 2			' Automatic
							oSheet.Rows(iRow).Cells(1, XL_RESULT).Interior.ColorIndex = 1			' Black
							oSheet.Rows(iRow).Cells(1, XL_RESULT).Interior.Pattern    = 1
							lTimerEnd = Timer
							iDuration = Round(lTimerEnd - lTimerStart)
							sDurationLog = "D: Execution time of test: " & iDuration & " s"
							
							WriteIntoLogAndOutputParamsCells oSheet.Rows(iRow), sDurationLog, ""
							
							If oSheet.Rows(iRow).RowHeight > 75 Then oSheet.Rows(iRow).RowHeight = 75
						
							oTraceLog.Message sDurationLog, LOG_SLOG
							' Jump out of row loop
							Exit Do
						ElseIf iResult = XL_DISPATCH_SKIP Then
							' Set result column
							oSheet.Rows(iRow).Cells(1, XL_RESULT).Value               = "skip"
							oSheet.Rows(iRow).Cells(1, XL_RESULT).Font.ColorIndex     = -4105		' Automatic
							oSheet.Rows(iRow).Cells(1, XL_RESULT).Interior.ColorIndex = 15			' 25% Gray
							oSheet.Rows(iRow).Cells(1, XL_RESULT).Interior.Pattern    = 1
						ElseIf iResult = XL_DISPATCH_PASS Then
							'Is FAIL expected? Then it's a Fail.  The cell will say pass but will be green for pass and red fail 
							if UCase(mid(trim(sExpectedResult),1,4)) = "FAIL" then
								bFailure = True
								oSheet.Rows(iRow).Cells(1, XL_RESULT).Value               = "pass(fail expected)"
								oTraceLog.Message "Passed (but fail expected) at: " & oSheet.name & "." & oSheet.Rows(iRow).Cells(1, XL_KEYWORD).Value, LOG_MESSAGE
								oSheet.Rows(iRow).Cells(1, XL_RESULT).Interior.ColorIndex = 3			' Red
								If bQCAutoDefect then 
									if bDefectThisSheet = False then 
										sDefectDescription = sDefectDescription & "Worksheet:" & oSheet.name & vbnewline  : oQRSDataObject.sQRSDefectDescription = sDefectDescription
										bDefectThisSheet = True
									end if 
									sDefectDescription = sDefectDescription & "Row:" & iRow & vbnewLine & RowToText(oSheet.Rows(iRow))  : oQRSDataObject.sQRSDefectDescription = sDefectDescription
									bDefectFound = True : oQRSDataObject.bQRSDefectFound = bDefectFound
								End If
							Else
								oSheet.Rows(iRow).Cells(1, XL_RESULT).Value               = "pass"
								oSheet.Rows(iRow).Cells(1, XL_RESULT).Interior.ColorIndex = 4			' Bright Green
							End if
							' Set result column
							oSheet.Rows(iRow).Cells(1, XL_RESULT).Font.ColorIndex     = -4105		' Automatic
							oSheet.Rows(iRow).Cells(1, XL_RESULT).Interior.Pattern    = 1
						ElseIf iResult = XL_DISPATCH_FAILCONTINUE Then
							' Set failure flag
							'bFailure = True
							'Is FAIL expected? Then it's a pass.  The cell will say fail but will be green for pass and red fail 
							If UCase(mid(trim(sExpectedResult),1,4)) = "FAIL" then
								oSheet.Rows(iRow).Cells(1, XL_RESULT).Value               = "fail(as expected)"
								oTraceLog.Message "Failed (as expected) at: " & oSheet.name & "." & oSheet.Rows(iRow).Cells(1, XL_KEYWORD).Value, LOG_MESSAGE
								oSheet.Rows(iRow).Cells(1, XL_RESULT).Interior.ColorIndex = 4			' Bright Green
							Else
								oSheet.Rows(iRow).Cells(1, XL_RESULT).Value               = "fail"
								oTraceLog.Message "Failed at: " & oSheet.name & "." & oSheet.Rows(iRow).Cells(1, XL_KEYWORD).Value, LOG_MESSAGE
								oSheet.Rows(iRow).Cells(1, XL_RESULT).Interior.ColorIndex = 3			' Red								
								If bQCAutoDefect then 
									If bDefectThisSheet = False then 
										sDefectDescription = sDefectDescription & "Worksheet:" & oSheet.name & vbnewline  : oQRSDataObject.sQRSDefectDescription = sDefectDescription
										bDefectThisSheet = True
									End If 
									sDefectDescription = sDefectDescription & "Row:" & iRow & vbnewLine & RowToText(oSheet.Rows(iRow))  : oQRSDataObject.sQRSDefectDescription = sDefectDescription
									bDefectFound = True : oQRSDataObject.bQRSDefectFound = bDefectFound
								End If
							End If
							' Set result column
							oSheet.Rows(iRow).Cells(1, XL_RESULT).Font.ColorIndex     = -4105		' Automatic
							oSheet.Rows(iRow).Cells(1, XL_RESULT).Interior.Pattern    = 1
						ElseIf iResult = XL_DISPATCH_FAIL Then
							' Set failure flag
							'Is FAIL expected? Then it's a pass.  The cell will say fail but will be green for pass and red fail 
							If UCase(mid(trim(sExpectedResult), 1, 4)) = "CONT" Then
								oSheet.Rows(iRow).Cells(1, XL_RESULT).Value               = "fail"
								oTraceLog.Message "Failed at: " & oSheet.name & "." & oSheet.Rows(iRow).Cells(1, XL_KEYWORD).Value, LOG_ERROR
								oSheet.Rows(iRow).Cells(1, XL_RESULT).Interior.ColorIndex = 3			' Red
								If bQCAutoDefect then 
									If bDefectThisSheet = False then 
										sDefectDescription = sDefectDescription & "Worksheet:" & oSheet.name & vbnewline  : oQRSDataObject.sQRSDefectDescription = sDefectDescription
										bDefectThisSheet = True
									End If 
									sDefectDescription = sDefectDescription & "Row:" & iRow & vbnewLine & RowToText(oSheet.Rows(iRow))  : oQRSDataObject.sQRSDefectDescription = sDefectDescription
									bDefectFound = True : oQRSDataObject.bQRSDefectFound = bDefectFound
								End If
								iResult = XL_DISPATCH_FAILCONTINUE
							ElseIf UCase(mid(trim(sExpectedResult),1,4)) = "FAIL" then
								oSheet.Rows(iRow).Cells(1, XL_RESULT).Value               = "fail(as expected)"
								oTraceLog.Message "Failed (as expected) at: " & oSheet.name & "." & oSheet.Rows(iRow).Cells(1, XL_KEYWORD).Value, LOG_MESSAGE
								oSheet.Rows(iRow).Cells(1, XL_RESULT).Interior.ColorIndex = 4			' Bright Green
							Else
								bFailure = True
								oSheet.Rows(iRow).Cells(1, XL_RESULT).Value               = "fail"
								oTraceLog.Message "Failed at: " & oSheet.name & "." & oSheet.Rows(iRow).Cells(1, XL_KEYWORD).Value, LOG_ERROR
								oSheet.Rows(iRow).Cells(1, XL_RESULT).Interior.ColorIndex = 3			' Red
								If bQCAutoDefect then 
									if bDefectThisSheet = False then 
										sDefectDescription = sDefectDescription & "Worksheet:" & oSheet.name & vbnewline  : oQRSDataObject.sQRSDefectDescription = sDefectDescription
										bDefectThisSheet = True
									end if 
									sDefectDescription = sDefectDescription & "Row:" & iRow & vbnewLine & RowToText(oSheet.Rows(iRow))  : oQRSDataObject.sQRSDefectDescription = sDefectDescription
									bDefectFound = True : oQRSDataObject.bQRSDefectFound = bDefectFound
								End If
							End If
							' Set result column
							oSheet.Rows(iRow).Cells(1, XL_RESULT).Font.ColorIndex     = -4105		' Automatic
							oSheet.Rows(iRow).Cells(1, XL_RESULT).Interior.Pattern    = 1
						ElseIf iResult = XL_DISPATCH_CANCEL Then
							' Set failure flag
							bFailure = True
		
							' Set result column
							oSheet.Rows(iRow).Cells(1, XL_RESULT).Value               = "cancel"
							oSheet.Rows(iRow).Cells(1, XL_RESULT).Font.ColorIndex     = -4105		' Automatic
							oSheet.Rows(iRow).Cells(1, XL_RESULT).Interior.ColorIndex = 3			' Red
							oSheet.Rows(iRow).Cells(1, XL_RESULT).Interior.Pattern    = 1
							oTraceLog.Message "Cancel", LOG_WARNING
							'If bQCAutoDefect then 
							'	if bDefectThisSheet = False then 
							'		sDefectDescription = sDefectDescription & "Worksheet:" & oSheet.name & vbnewline  : oQRSDataObject.sQRSDefectDescription = sDefectDescription
							'		bDefectThisSheet = True
							'	end if 
							'	sDefectDescription = sDefectDescription & "Row:" & iRow & vbnewLine & RowToText(oSheet.Rows(iRow))  : oQRSDataObject.sQRSDefectDescription = sDefectDescription
							'	bDefectFound = True : oQRSDataObject.bQRSDefectFound = bDefectFound
							'End If
						Else
	
							' Set failure flag
							bFailure = True
	
							' Set result column
							oSheet.Rows(iRow).Cells(1, XL_RESULT).Value               = "???"
							oSheet.Rows(iRow).Cells(1, XL_RESULT).Font.ColorIndex     = -4105		' Automatic
							oSheet.Rows(iRow).Cells(1, XL_RESULT).Interior.ColorIndex = 6			' Yellow
							oSheet.Rows(iRow).Cells(1, XL_RESULT).Interior.Pattern    = 1
							oTraceLog.Message "Unknown state", LOG_WARNING
							If bQCAutoDefect then 
								if bDefectThisSheet = False then 
									sDefectDescription = sDefectDescription & "Worksheet:" & oSheet.name & vbnewline  : oQRSDataObject.sQRSDefectDescription = sDefectDescription
									bDefectThisSheet = True
								end if 
								sDefectDescription = sDefectDescription & "Row:" & iRow & vbnewLine & RowToText(oSheet.Rows(iRow))  : oQRSDataObject.sQRSDefectDescription = sDefectDescription
								bDefectFound = True : oQRSDataObject.bQRSDefectFound = bDefectFound
							End If
						End If
						
						'what it is?
						sParm = " - " & oSheet.Rows(iRow).cells(1, XL_PARM_001)
						for i = XL_PARM_002 to XL_PARM_010
							sParm = sParm & "," & oSheet.Rows(iRow).cells(1, i)
						Next
						
						if bAnimate then oSheet.Range("A" & iRow & ":F" & iRow).show 'JE added
		
						'Display Alerts here for 'PASS' or FAILS
						If bLogLevel Then 
							If iResult = XL_DISPATCH_PASS Or iResult = XL_DISPATCH_FAIL Then
								sGridTitles = "Work Sheet Name|"
								sGridTitles = sGridTitles&oSheet.Rows(1).Cells(1, XL_RESULT).Value & "|"
								sGridTitles = sGridTitles&oSheet.Rows(1).Cells(1, XL_KEYWORD).Value & "|"
								sGridTitles = sGridTitles&oSheet.Rows(1).Cells(1, XL_LOG).Value & " "
	
								sGridValues=oSheet.name & "|"
								if	iResult=XL_DISPATCH_FAIL then
									sGridValues=sGridValues & "fail|"
								else
									sGridValues=sGridValues & "pass|"
								end if
								sGridValues = sGridValues&oSheet.Rows(iRow).Cells(1, XL_KEYWORD).Value & "|"
								sGridValues = sGridValues&oSheet.Rows(iRow).Cells(1, XL_LOG).Value & " "
								'Grep just the current spreadsheet name
								aWorkbookName = Split(sWorkbook, "\", -1)
								oXlApp.Visible = True'
								oSheet.Activate
								oSheet.Range("A" & iRow & ":F" & iRow).show
								LogAlert aWorkbookName(ubound (aWorkbookName)), sGridTitles, sGridValues
								If iResult = XL_DISPATCH_PASS Then
									WaitTime (300)
								Else
									WaitTime (500)'Display failures longer
								End if
								'WaitTime (ShowAlert (aWorkbookName(ubound (aWorkbookName)),sGridTitles,sGridValues))
								oXlApp.Visible = False
							End If
						end if 'bLogLevel
						
						' Handle test failure and errors
						If bFailure Then
							' Failure/error causes current worksheet to end

							If dictStreams.Count = 0 Then
								'if streams are not used, jump out of row loop
								
								Exit Do
							Else
								'if streams are used, set actual row streams as failed first
								arrStreams = Split(oSheet.Rows(iRow).Cells(1, XL_STREAM).Text, ",")
								If UBound(arrStreams) = -1 Then
									'set all streams as failed
									For Each sKey In dictStreams.Keys
										
										WriteIntoLogAndOutputParamsCells oSheet.Rows(iRow), "Stream failed: " & sKey, ""
										dictStreams(sKey) = "F"
									Next
								Else
									For Each sKey In arrStreams
										WriteIntoLogAndOutputParamsCells oSheet.Rows(iRow), "Stream failed: " & sKey, ""
										dictStreams(sKey) = "F"
									Next
								End If
							
								iCounter = 0
								'exit test if all streams failed or set to skip
								For Each sKey In dictStreams.Keys
									If dictStreams(sKey) = "F" Or dictStreams(sKey) = "S" Then
										iCounter = iCounter + 1
									End If
								Next
								'MsgBox "Failed streams: " & iCounter & vblf & " / " & dictStreams.Count
								If iCounter = dictStreams.Count Then
									WriteIntoLogAndOutputParamsCells oSheet.Rows(iRow), "Execution is over: all defined workbook streams failed or set to skip!", ""
									Exit Do
								End If
							End If
						End If
					
						iRow = iRow + 1
					Loop
						
					' Save workbook at end of each worksheet
					iResult = SaveWorkbook(oWorkbook)
	
					' Handle test failure and errors
					If bFailure Then
						' Failure/error causes current workbook to end
						' Jump out of sheet loop					
						
						lTimerEnd = Timer
						iDuration = Round(lTimerEnd - lTimerStart)
						sDurationLog = "D: Execution time of test: " & iDuration & " s" 
						'REM have to find where is end
						For i = iRow To oSheet.UsedRange.Rows.Count
							If UCase(oSheet.Rows(i).Cells(1, XL_KEYWORD)) = "END" Then
								'we dont know where the end is, then we write duration here
								WriteIntoLogAndOutputParamsCells oSheet.Rows(i), sDurationLog, ""
								
								If oSheet.Rows(i).RowHeight > 75 Then 
									oSheet.Rows(i).RowHeight = 75
								End If	
								Exit For
							End If						
						Next
											
						oTraceLog.Message sDurationLog, LOG_SLOG
						
						If oQC.IsQCRun Then												
							'set test result to ALM
							globCurrentRun.Status = "Failed"													
						End If										
						
						iResult = SaveWorkbook(oWorkbook)
						
						Set oSheet = Nothing
						Exit For
					End If
					
				Next
				
				Set oSheet = Nothing
			Else
				oTraceLog.Message "Failed at spreadsheet validation.", LOG_ERROR
				oTraceLog.Message "Test Failed", LOG_ERROR
				If oQC.IsQCRun Then												
					'set test result to ALM
					globCurrentRun.Status = "Failed"
						
					oTraceLog.StepMessage "VALIDATION PROBLEM", XL_DISPATCH_FAIL, "Failed at spreadsheet validation.", "", "" 
				End if
				SaveWorkbook(oWorkbook)
			End If
		Else
			oTraceLog.Message "Test is not in status Active.", LOG_ERROR
			globCurrentRun.Status = "Failed"
			oTraceLog.StepMessage "Test is not in status Active.", XL_DISPATCH_FAIL, "", "", ""		
			SaveWorkbook(oWorkbook)
		End If	'bContinue
			

		'TODO: is it OK to set iResult here?
		' Close workbook after all worksheets
		iResult = CloseWorkbook(oWorkbook)
		
		Set oWorkbook = Nothing
		'if test is run from QC, then upload results as attachment to QC Current Run 
		 If oQC.IsQCRun() then 		 	
			 oQC.UploadAttachmentToQCRun(sWorkbook)
			 If bPerformance Then
				 oQC.UploadAttachmentToQCRun(sPerformanceFile)
			 End If
			 if bQCAutoDefect AND bDefectFound then
				 'Create defect
			    Set oBug = oQC.AddDefect(sQCAssignee, globTDConnection.UserName, "Automated Run Workbook: " & oFso.GetFileName(sWorkbook), sDefectDescription, sWorkbook)
				 'Link the defect to the current run
				 globCurrentRun.BugLinkFactory.AddItem(oBug)
				 
				 'Send mail 
				 oQC.MailDefect oBug.ID, sQCEmailRecipients, "", _
							 "QC Defect - Domain:" & globTDConnection.DomainName & ", Project:" & globTDConnection.ProjectName & ", Automated Run Workbook: " & oFso.GetFileName(sWorkbook), _
							 "See the description below for details of sheets and rows."
				 Set oBug = Nothing
			 End If
			 'globCurrentRun.Status = "NO_RUN"
			 'globCurrentRun.Post			
			 'ExitTest
		 End If
		
	
		' Close Excel application
		oXlApp.WindowState=-4137	'Maximize window		
	
		oXlApp.Quit
		Set oXlApp = Nothing
		
		'capture screen before killing processes			
		If bFailure = True Then	
			'Commented the below out because TAUtilities is currently failing (activeX issue or permissions)
			'Screenshot globResultPath & "\" & "screenshot@failure_" & YYYYMMDDHHMMSS(Now()) & ".png", True
		End If
		
	'	WaitTime(10000)
	'	bAutoCleanUp=False
		If bAutoCleanUp Then		
			WindowsProcessesCleanUp dFrameworkStart, bFailure
		End If
		
		oTraceLog.Message "MAIN: " & "Test finished for workbook " & sWorkbook, LOG_MESSAGE
		
		
		Select Case globRetVal
			Case XL_DISPATCH_FAIL
				oVBSFramework.oTraceLog.HostLogMessage "Result: Fail"
			Case XL_DISPATCH_PASS
				oVBSFramework.oTraceLog.HostLogMessage "Result: Pass"
			Case XL_DISPATCH_CANCEL
				oVBSFramework.oTraceLog.HostLogMessage "Result: Cancel"
			Case XL_DISPATCH_FAILCONTINUE
				oVBSFramework.oTraceLog.HostLogMessage "Result: Fail"
			Case XL_DISPATCH_UNKNOWN
				oVBSFramework.oTraceLog.HostLogMessage "Result: Unknown"
			Case Else
				oVBSFramework.oTraceLog.HostLogMessage "Result: Unknown"
		End Select 
		
		sXUnitFileText = sXUnitFileText & _
						"</testsuite>" & vbNewLine & _
						"</testsuites>" 

		'sXUnitTestClassName = replace(sXUnitTestClassName, """", "")						
'msgbox Left(globResultPath, Len(globResultPath)) & ".xml" & vbNewLine & sXUnitFileText
		Call WriteAllFileText(Left(globResultPath, Len(globResultPath)) & ".xml", sXUnitFileText)

		oVBSFramework.oTraceLog.HostLogMessage "Duration: " & iDuration & " s"
		oTraceLog.Message "END", LOG_MESSAGE
		oTraceLog.HostLogMessage "Main: " & "Test finished for workbook " & sWorkbook 
		oTraceLog.HostLogMessage "END "
		'oTraceLog.HostLogMessage " END " & vbNewLine & _
					'"**** * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * ****"												
	

	End Sub

	'==============================================================================================
	' Function:		ShowAlert
	' Purpose:		To display desktop alerts of test progress.
	'
	' Parameters:	Main title,grid title and grid values
	'				Requires Indicator.dll to be registered on client pc
	' Returns:		Time in milliseconds for which the alet is displayed
	'
	'==============================================================================================
	Private Function ShowAlert(sLabel, sGridTitle, sGridValues)
		Dim oShell, sCommand, sCurrentDir, iInterval
		Dim sFileStream, sText
		On error resume next

		Set oShell = CreateObject ("WSCript.Shell")

		sCommand =  "GPSPAlert.vbs"'vb script to show alert

		'if GPSAlert.vbs does not exist in the current directory then build it
		'normally this will be the QTP installation directory.

		sCurrentDir=oShell.CurrentDirectory&"\"
		If not oFso.FileExists(sCurrentDir&sCommand) Then
			'lets build vbscript to execute
			Set sFileStream = oFso.OpenTextFile(sCommand,2, True,False)'2 is for new file

			sFileStream.WriteLine("Set indicator = CreateObject("&chr(34)&"indicator.myInd"&chr(34)&")")
			sFileStream.WriteLine("If WScript.Arguments.Named.Exists("&chr(34)&"label"&chr(34)&") Then")
			sFileStream.WriteLine(chr(9)&"indicator.Info1 = WScript.Arguments.Named("&chr(34)&"label"&chr(34)&")")
			sFileStream.WriteLine("End if")

			sFileStream.WriteLine("If WScript.Arguments.Named.Exists("&chr(34)&"area"&chr(34)&") Then")
			sFileStream.WriteLine(chr(9)&"indicator.Area = WScript.Arguments.Named("&chr(34)&"area"&chr(34)&")")
			sFileStream.WriteLine("End if")

			sFileStream.WriteLine("If WScript.Arguments.Named.Exists("&chr(34)&"interval"&chr(34)&") Then")
			sFileStream.WriteLine(chr(9)&"indicator.Interval = WScript.Arguments.Named("&chr(34)&"interval"&chr(34)&")")
			sFileStream.WriteLine("End if")

			sFileStream.WriteLine("If WScript.Arguments.Named.Exists("&chr(34)&"gridTitle"&chr(34)&") Then")
			sFileStream.WriteLine(chr(9)&"indicator.gridTitle = WScript.Arguments.Named("&chr(34)&"gridTitle"&chr(34)&")")
			sFileStream.WriteLine("End if")

			sFileStream.WriteLine("If WScript.Arguments.Named.Exists("&chr(34)&"gridValues"&chr(34)&") Then")
			sFileStream.WriteLine(chr(9)&"indicator.gridValues = WScript.Arguments.Named("&chr(34)&"gridValues"&chr(34)&")")
			sFileStream.WriteLine("End if")

			sFileStream.WriteLine("indicator.show")
			sFileStream.WriteLine("Set indicator = Nothing")
			sFileStream.WriteLine("WScript.Quit")

			sFileStream.Close

			Set sFileStream =Nothing
		End If
		sCommand = sCommand & " /area:" &"6"'display at top middle
		iInterval=3.5'seconds
		sCommand = sCommand & " /interval:" &CStr(iInterval*1000)'in milliseconds
		sCommand = sCommand & " /label:" & chr(34)&sLabel&chr(34)
		sCommand = sCommand & " /gridTitle:" & chr(34)&sGridTitle&chr(34)
		sCommand = sCommand & " /gridValues:" & chr(34)&sGridValues&chr(34)
		oShell.Run sCommand
		Set oShell = Nothing

		If Err.Number <> 0 Then
			ShowAlert=0
			msgbox(sCommand&" file missing")
			Exit Function
		End if
		ShowAlert=iInterval
	End Function

	'==============================================================================================
	' Function:		LogAlert
	' Purpose:		To display desktop alerts of test progress.
	'
	' Parameters:	Main title,grid title and grid values
	'				Requires Indicator.dll to be registered on client pc
	' Returns:		True or False
	'
	'==============================================================================================
	Private Function  LogAlert(sLabel, sGridTitle, sGridValues)
		Dim sFileStream, sText
		On error resume next

		'update logfile
		Set sFileStream = oFso.OpenTextFile(sVBSFrameworkDir&"\"&sLogFileName,8, True,False)'8 is for appending to file
		sFileStream.WriteLine(sLabel)
		sFileStream.WriteLine(sGridTitle)
		sFileStream.WriteLine(sGridValues)
		sFileStream.Close
		Set sFileStream =Nothing
		Set oFso = Nothing
		LogAlert=true
	End Function
	

'	==============================================================================================
	Public Function IsQTPAddInInstalled(sAddInName)
		Dim i, oQTP, bAddInExists
	
		Set oQTP = CreateObject("QuickTest.Application")
		bAddInExists = False

		For i = 1 To oQTP.AddIns.Count
			MsgBox "AddIn #" & i & vblf & "AddIn Name: " & myQTP.AddIns.Item(i).Name & vblf & "AddIn Status: " & myQTP.AddIns.Item(i).status & vblf & "AddIn Version: " & myQTP.AddIns.Item(i).Version
			If oQTP.AddIns.Item(i).Name = sAddInName Then
				bAddInExists = True
				Exit For
			End If
		Next
	
		IsQTPAddInInstalled = bAddInExists
	End Function

	'==============================================================================================
	Public Function IsQTPAddInLoaded(sAddInName)
	   Dim i, oQTP, bAddInLoaded
	
		Set oQTP = CreateObject("QuickTest.Application")
		bAddInLoaded = False
	
		For i = 1 To oQTP.AddIns.Count
			If oQTP.AddIns.Item(i).Name = sAddInName Then
				If oQTP.AddIns.Item(i).Status = "Active" Then
					bAddInLoaded = True
					Exit For
				End If
			End If
		Next
	
		IsQTPAddInLoaded = bAddInLoaded
	End Function
	
	'==============================================================================================
	' Function:
	'==============================================================================================
	Sub WindowsProcessesCleanUp(dFrameworkStart, bFailure)
		Dim oWMIDateTime, oWMIService, colProcessList, colProcessList2, oProcess, oProcess2, iCount, sUserName, sDomain, sScreenshotFullPath, iProcessID
		Dim oTestObject
				
		
		'construct datetime in WMI format
		Set oWMIDateTime = CreateObject("WbemScripting.SWbemDateTime")
		oWMIDateTime.SetVarDate(dFrameworkStart)
		
		Set oWMIService = GetObject("winmgmts:{impersonationLevel=impersonate}!\\" & "." & "\root\cimv2")
		Set colProcessList = oWMIService.ExecQuery("Select * from Win32_Process WHERE CreationDate > '" & oWMIDateTime & "'")
		
		On Error Resume Next
		iCount = colProcessList.Count
		If (Err.Number <> 0) Then
		
			oTraceLog.Message "WindowsProcessesCleanUp(), WMI error FIRST:" & Err.Description,LOG_ERROR 
			Err.Clear
		Else
			If globDebug Then
				oTraceLog.Message "Looping processes", LOG_MESSAGE 
			End If
			
			For Each oProcess in colProcessList
				iProcessID = oProcess.ProcessID 
			'	MsgBox "Proces name: "& oProcess.Name & ",Proces id: " & iProcessID & ", Parent ID: " & oProcess.ParentProcessId	
				If globDebug Then
					oTraceLog.Message "Proces name: "& oProcess.Name & ",Proces id: " & iProcessID & ", Parent ID: " & oProcess.ParentProcessId , LOG_MESSAGE 
				End if
				Select Case UCase(oProcess.Name)
					Case "ARCHIE32.EXE"
						'specific handling for DREP application					
						
						Set oTestObject = oVBSFramework.oTestObjects.Item("DREP")
						oTestObject.StopDREP Null, ""
						Set oTestObject = Nothing
						'terminate process if still exists
						WaitTime(1000)
						Set colProcessList2 = oWMIService.ExecQuery("Select * from Win32_Process WHERE ProcessID = " & iProcessID)
						For Each oProcess2 In colProcessList2
							oProcess2.Terminate
						Next
						Set colProcessList2 = Nothing
						oTraceLog.Message "WindowsProcessesCleanUp(): Closing DREP", LOG_MESSAGE
					Case "MASTER.EXE"
						'specific handling for ENDUR application					
						
						Set oTestObject = oVBSFramework.oTestObjects.Item("ENDUR")
						oTestObject.StopEndur Null, ""
						Set oTestObject = Nothing
						'terminate process if still exists
						WaitTime(1000)
						Set colProcessList2 = oWMIService.ExecQuery("Select * from Win32_Process WHERE ProcessID = " & iProcessID)
						For Each oProcess2 In colProcessList2
							oProcess2.Terminate
						Next
						Set colProcessList2 = Nothing
						oTraceLog.Message"WindowsProcessesCleanUp(): Closing ENDUR", LOG_MESSAGE
'					Case "SAPLOGON.EXE"
						'specific handling for SAP application
'						REM Reporter.ReportEvent micDone, "WindowsProcessesCleanUp()", "Closing SAP Session"
'						Set oTestObject = oVBSFramework.oTestObjects.Item("SAP")
'						oTestObject.StopSAP ""
'						Set oTestObject = Nothing
						'terminate process if still exists
'						WaitTime 1
'						Set colProcessList2 = oWMIService.ExecQuery("Select * from Win32_Process WHERE ProcessID = " & iProcessID)
'						For Each oProcess2 In colProcessList2
'							oProcess2.Terminate
'						Next
'						Set colProcessList2 = Nothing
					Case "WMIPRVSE.EXE", "RVD.EXE", "QTAUTOMATIONAGENT.EXE", "SAPLOGON.EXE","CSCRIPT.EXE"
						'ignore	
					Case "EXCEL.EXE"						
						If iProcessID <> CInt(Trim(sExcelProcesID)) Then
							oProcess.GetOwner sUserName, sDomain						
							oProcess.Terminate()
							oTraceLog.Message "WindowsProcessesCleanUp(): Terminating process [" & oProcess.Name & "] owned by " & sDomain & "\" & sUserName & ", with PID: " & iProcessID , LOG_MESSAGE	
					
						End If		
					Case Else
						'force terminate process
						oProcess.GetOwner sUserName, sDomain						
						oProcess.Terminate()
						oTraceLog.Message "WindowsProcessesCleanUp(): Terminating process [" & oProcess.Name & "] owned by " & sDomain & "\" & sUserName& ",PID: " & iProcessID , LOG_MESSAGE
						
				End Select
			Next
		End If
		
		'SAP handling
		Set colProcessList = oWMIService.ExecQuery("Select * from Win32_Process WHERE name = 'saplogon.exe'")
		If (Err.Number <> 0) Then
			Rem Reporter.ReportEvent 3, "WindowsProcessesCleanUp()", "WMI error:" & Err.Description
			oTraceLog.Message "WindowsProcessesCleanUp(): WMI error:" & Err.Description, LOG_ERROR
			Err.Clear
		Else
			For Each oProcess in colProcessList
				Rem Reporter.ReportEvent micDone, "WindowsProcessesCleanUp()", "Closing SAP Session"
				oTraceLog.Message "WindowsProcessesCleanUp(): Closing SAP Session", LOG_MESSAGE
				Set oTestObject = oVBSFramework.oTestObjects.Item("SAP")
				oTestObject.StopSAP ""
				Set oTestObject = Nothing
			Next
		End If
		On Error Goto 0
		
		'capture screen after killing processes
		If bFailure = True Then
		'	Screenshot PATH_RESOURCES & "screenshot@failure_" & YYYYMMDDHHMMSS(Now()) & "_cleanup.png", False
		
			'Commented the below out because TAUtilities is currently failing (activeX issue or permissions)
			'Screenshot globResultPath & "\" & "screenshot@failure_" & YYYYMMDDHHMMSS(Now()) & "_cleanup.png", False
		End If
		
		'Set dictWhiteList = Nothing
		Set colProcessList = Nothing
		Set oWMIService = Nothing
		Set oWMIDateTime = Nothing	
	End Sub

	'==============================================================================================
	' Function:     CopyWorkbook
	' Purpose:      Copy workbook to Test Results directory
	' Parameters:   Source workbook,Target workbook
	' Returns:      path+filename of workbok in result directory
	'==============================================================================================
	Public Function CopyWorkbook(sSourceWB, sTargetWB)		

		'sTargetWB = sTargetPath & oFso.GetBaseName(sSourceWB) & "_RESULT" & YYYYMMDDHHMM(Now()) & "." & oFso.GetExtensionName(sSourceWB)
		'copy workbook in given target
		oFso.CopyFile sSourceWB, sTargetWB 'TODO add error handling, and tracelog

		' Return path+filename of workbook in result directory
		CopyWorkbook = sTargetWB
	End Function
	
	'==============================================================================================
	' Function:     GetResultWorkBookPath
	' Purpose:      Create result file name
	' Parameters:   Source workbook, Target path
	' Returns:      path+filename of workbok in result directory
	' Author: 		
	'==============================================================================================
	Public Function GetResultWorkBookPath(sSourceWB, sTargetPath)
		Dim sTargetWB	
			
		On Error Resume Next 'For when the results folder already exists
	 		oFso.createfolder (sTargetPath)
		On Error GoTo 0
		
		sTargetWB = sTargetPath & oFso.GetBaseName(sSourceWB) & "_RESULT" & YYYYMMDDHHMM(globTestStart) & "." & oFso.GetExtensionName(sSourceWB)
		' Return path+filename of workbook in result directory
		GetResultWorkBookPath = sTargetWB
	End Function
	
	'==============================================================================================
	Private Sub RemoveFormulasKeepValuesInWorkbook(sWorkbook)
		Dim oWorkbook
		
		Set oWorkbook = oXlApp.Workbooks.Open(sWorkbook)
		
		oXlApp.CalculateFull
		If globDebug Then
			oTraceLog.Message "Removing formulas and keeping values in workbook", LOG_DEBUG
		End If
		'TODO: log information about replacements
		oWorkbook.Worksheets(1).UsedRange.Copy
		oWorkbook.Worksheets(1).UsedRange.PasteSpecial -4163
		
		oWorkbook.Worksheets(1).Cells(1, 1).Select
		oWorkbook.Save
		oWorkbook.Close
		Set oWorkbook = Nothing
	End Sub
	
	


	'==============================================================================================
	' Function:		OpenWorkbook
	' Purpose:		Open the specified workbook
	'
	' Parameters:	Source workbook
	'
	' Returns:		Workbook object for the opened workbook
	'
	'==============================================================================================
	Private Function OpenWorkbook(sWorkbook)
		Dim iResult
		
		On Error Resume Next
		
		Set OpenWorkbook = oXlApp.Workbooks.Open(sWorkbook)
		
		If Err.Number <> 0 Then
			oTraceLog.Message "Problem with openning workbook: " & Err.Description, LOG_ERROR
		End If
		On Error Goto 0
	End Function

	'==============================================================================================
	' Function:		SaveWorkbook
	' Purpose:		Save the specified workbook
	'
	' Parameters:	object Workbook
	'
	' Returns:		XL_DISPATCH_PASS or XL_DISPATCH_FAIL
	'
	'==============================================================================================
	Private Function SaveWorkbook(oWorkbook)
		Dim iRetVal

		oWorkbook.Save

		' TODO: Determine if save was successful
		iRetVal = XL_DISPATCH_PASS

		SaveWorkbook = iRetVal
	End Function

	'==============================================================================================
	' Function:		CloseWorkbook
	' Purpose:		Close the speficied workbook
	'
	' Parameters:	object Workbook
	'
	' Returns:		XL_DISPATCH_PASS or XL_DISPATCH_FAIL
	'
	'==============================================================================================
	Private Function CloseWorkbook(oWorkbook)
		Dim iRetVal

		' Close workbook
		oWorkbook.Close
		
		' TODO: Determine if close was successful
		iRetVal = XL_DISPATCH_PASS

		CloseWorkbook = iRetVal
	End Function

	'==============================================================================================
	'
	'==============================================================================================
	Function RowToText(oRow)
		Dim i
		Dim sRetVal

		sRetVal = "result=" & oRow.Cells(1, XL_RESULT).Text & vbnewline & _
					"keyword=" & oRow.Cells(1, XL_KEYWORD).Text & vbnewline & _
					"comment=" & oRow.Cells(1, XL_COMMENT).Text & vbnewline & _
					"output=" & oRow.Cells(1, XL_LOG).Text & vbnewline
					
		For i = XL_PARM_001 to XL_PARM_010
			If oRow.Cells(1, i).Text = "" Then Exit For
			sRetVal = sRetVal & "parm_" & Left("00" & i - XL_PARM_001 + 1, 3) & "=" & oRow.Cells(1, i).Text & vbnewline
		Next

		sRetVal = sRetVal & "message=" & oRow.Cells(1, XL_OUTPUT_PARAMS).Text & vbnewline
		RowToText = sRetVal
	End Function
	
	'==============================================================================================
	'
	'==============================================================================================
	'Public Function IsQCRun()
	'	Dim bRetVal
	'	
	'	bRetVal = False
	'	
	'	If globTDConnection.Connected then 
	'		If Not globCurrentTSTest.Test Is Nothing Then
	'			bRetVal = True
	'		End If
	'	End If
	'	
	'	IsQCRun = bRetVal
	'End Function
	
	'==============================================================================================
	'
	'==============================================================================================
	'Public Function UploadAttachmentToQCRun(sFullPath)
	'	Dim oAttachFact, oAttachment
	'
	'	Set oAttachFact = globCurrentRun.Attachments
	'	Set oAttachment = oAttachFact.AddItem(Null)
	'	oAttachment.FileName = sFullPath
	'	oAttachment.Type = 1
	'	oAttachment.Post
	'	oAttachment.Refresh
	'	
	'	Set oAttachment = Nothing
	'	Set oAttachFact = Nothing
	'End Function

	'==============================================================================================
	'This wrapper is used in the launch script and didin't want to create a QC object just for that
	'==============================================================================================
	Public Function GetAttachmentFileFromQC(sPrefix, sExtension)
		'Dim oQC
		
	'	CreateFolderFromPath(PATH_RESOURCES)
		CreateFolderFromPath(globResultPath)
		'Set oQC = New clsQC
	'	GetAttachmentFileFromQC = oQC.GetAttachmentFileFromQC(sPrefix, sExtension, PATH_RESOURCES)
		GetAttachmentFileFromQC = oQC.GetAttachmentFileFromQC(sPrefix, sExtension, globResultPath)
		'Set oQC = Nothing
	End Function

	'==============================================================================================
	'This function will get the test spreadsheet from QC using the following rules:
	'If no xls is present then write an error to the QTP results
	'If only one XLS file is present then use it as the test XLS
	'If more than one is present then use the one that matches the QC test name
	'If more than one xls is present and none matches the name of the QC test then write an error to the QTP results.
	'In the future consider a "fuzzy" match to use the xls closest to the QC test name.
	'==============================================================================================
	Public Function GetTestSpreadsheetFromQC()		
		Dim sTestSpreadsheet
		Dim sMessage
		Dim iRetVal
		
		Const sRoutine = "clsVBSFramework.GetTestSpreadsheetFromQC"
		oTraceLog.HostLogMessage("> ENTERED: " & sRoutine)
		
		CreateFolderFromPath(PATH_RESOURCES)

		sTestSpreadsheet = ""

		If oQC.IsQCRun() Then		
			'Check how many spreadsheets attached
		
			Select Case oQC.CountAttachedFilesByPattern(".*\.xls")
			Case 0		'No spreadsheet
				sMessage = "GetTestSpreadsheetFromQC(): No XLS file attached."			
				iRetVal = XL_DISPATCH_FAIL
			Case 1		'1 spreadsheet so use it
				sMessage = "GetTestSpreadsheetFromQC(): One XLS file attached. Will use this as the test spreadsheet."			
				iRetVal = XL_DISPATCH_PASS
				sTestSpreadsheet = oQC.GetAttachmentFileFromQCByPattern(".*\.xls", PATH_RESOURCES)	
				If sTestSpreadsheet = "" Then
					iRetVal = 	XL_DISPATCH_FAIL
					sMessage = "Cannot get attachment. Look into Trace.Log"
				End If		
			Case Else	'More than one spreadsheet
				'Try to find one that matches the test name
				sMessage = "GetTestSpreadsheetFromQC(): More than one XLS file attached. Looking for one that matches the current test name."
			
				sTestSpreadsheet = oQC.GetAttachmentFileFromQCByPattern(globCurrentTSTest.Test.Name &  "\.xls", PATH_RESOURCES)
				If sTestSpreadsheet = "" Then
					sMessage = "GetTestSpreadsheetFromQC(): No attached XLS file matches the current test name."				
					
					iRetVal = XL_DISPATCH_FAIL
				Else
					sMessage = "GetTestSpreadsheetFromQC(): Found attached XLS file matching the current test name. Will use this as the test spreadsheet."				
					
					iRetVal = XL_DISPATCH_PASS
				End If 
			End Select
			
			If iRetVal = XL_DISPATCH_FAIL Then
				oTraceLog.HostLogMessage sMessage
			Else
				oTraceLog.HostLogMessage sMessage
			End If	
			
			oTraceLog.StepMessage "Get Test Spreadsheet From QC", iRetVal, sMessage, , ""
		Else 
			oTraceLog.HostLogMessage "QC not running, test is performed probably from launcher, please check your running option"	
			oTraceLog.StepMessage "QC connection problem", XL_DISPATCH_FAIL, "", , ""
			globCurrentRun.Status = "Failed"
			globRetVal = XL_DISPATCH_FAIL
			GetTestSpreadsheetFromQC = ""			
			oTraceLog.HostLogMessage("< EXITED: " & sRoutine): Exit Function
		End If
				
		GetTestSpreadsheetFromQC = sTestSpreadsheet
		oTraceLog.HostLogMessage("< EXITED: " & sRoutine)
	End Function

	'==============================================================================================
	'
	'==============================================================================================
	Private Sub CreateFolderFromPath(sFolder)
		Dim sParentFolder
	
		sParentFolder = oFso.GetParentFolderName(sFolder)
		If Not oFso.FolderExists(sFolder) Then
			If Not oFso.FolderExists(sParentFolder) Then
				CreateFolderFromPath(sParentFolder)
			End if
			oFso.CreateFolder(sFolder)
		End If
	
	End Sub

	'==============================================================================================
	'
	'==============================================================================================
	'Public Function GetAttachmentFileFromQC(sPrefix, sExtension)
	'	Dim oCurrentTest, oAttachFact, oAttachment
	'	Dim AttachList, sAttachList, sFilename
	'
	'	Set oCurrentTest = globCurrentTSTest.Test 
	'	Set oAttachFact = oCurrentTest.Attachments
	'	Set AttachList = oAttachFact.NewList("")
	'
	'	sFilename = ""
	'	If Attachlist.Count > 0 Then
	'		For Each oAttachment In AttachList
	'			If UCase(Left(oAttachment.Name(1), Len(sPrefix))) = uCase(sPrefix) And UCase(Right(oAttachment.Name(1), Len(sExtension))) = UCase(sExtension) Then
	'				'Copy the file to the specified path
	'				oAttachment.Load True, ""
	'				sFilename = oAttachment.FileName
	'				Exit For
	'			End If
	'		Next
	'	Else
	'		'No attachment found
	'	End If
	'
	'	Set oCurrentTest = Nothing
	'	Set oAttachFact = Nothing
	'	Set AttachList = Nothing
	'
	'	GetAttachmentFileFromQC = sFilename
	'
	'End Function

	'==============================================================================================
	'
	'==============================================================================================
	'Public Function AddDefect(sAssignedTo, sDetectedBy, sSummary, sDescription, sAttachmentFullPath)
	'	Dim oQCConnection
	'	Dim oBugFact, oBug
	'	Dim oAttachFact, oAttachment
	'
	'	Set oQCConnection = globTDConnection 
	'	Set oBugFact = oQCConnection.BugFactory
	'
	'	Set oBug = oBugFact.AddItem(Null)
	'
	'	oBug.AutoPost = False
	'	oBug.AssignedTo = sAssignedTo
	'	oBug.DetectedBy = sDetectedBy
	'	oBug.Priority = "1-Low"
	'	oBug.Status = "New"
	'	oBug.Summary = sSummary
	'	oBug.Field("BG_DESCRIPTION") = sDescription
	'	oBug.Field("BG_DETECTION_DATE") = now()
	'	oBug.Field("BG_SEVERITY") = "2-Medium"
	'	oBug.Field("BG_REPRODUCIBLE") = "Y"
	'	oBug.Field("BG_PRIORITY") = "2-Medium"
	'	oBug.Field("BG_RESPONSIBLE") = sAssignedTo
	'	oBug.Field("BG_STATUS") = "New"
	'	'oBug.Field("BG_CATEGORY") = "Automated"
	'	oBug.Field("BG_PROJECT") = globTDConnection.DomainName & "." & globTDConnection.ProjectName
	'	oBug.Field("BG_USER_01") = "1abcd"
	'	oBug.Field("BG_USER_03") = "English"	'Language
	'	oBug.Field("BG_USER_04") = "N"			'Regression
	'	oBug.Field("BG_USER_05") = "Test Automation"	'Category
	'
	'	oBug.Post
	'	oBug.Refresh
	'	
	'	'Now Attach file
	'	If sAttachmentFullPath <> "" then 
	'		Set oAttachFact = oBug.Attachments
	'		Set oAttachment = oAttachFact.AddItem(Null)
	'		oAttachment.FileName = sAttachmentFullPath
	'		oAttachment.Type = 1
	'		oAttachment.Post
	'		oAttachment.Refresh
	'
	'		Set oAttachFact = Nothing
	'	End If
	'
	'	Set AddDefect = oBug
	'
	' End Function

	'==============================================================================================
	'
	'==============================================================================================
	'Sub MailDefect(iBugID, sMailTo, sMailCC, sMailSubject, sMailComment)
	'	Dim i
	'	Dim oQCConnection
	'	Dim oBugFact, oBug, oBugList
	'
	'	Set oQCConnection = globTDConnection 
	'	Set oBugFact = oQCConnection.BugFactory
	'	Set oBugList = oBugFact.newlist("")
	'
	'	For i = oBugList.count to 1 step -1
	'		Set oBug = oBuglist.item(i)
	'		If oBuglist.item(i).ID = iBugID Then
	'			Exit For 
	'		End If
	'	Next
	'
	'	oBug.Mail sMailTo, sMailCC, 4, sMailSubject, sMailComment
	'
	'End Sub
 	
	'==============================================================================================
	' Function:		K_Delay
	' Purpose:		Delays script execution for specified number of seconds
	'
	' Parameters:	Parm001 - Duration - number of seconds to WaitTime
	'
	' Returns:		TRUE  - if delay successful
	'				FALSE - if delay fails
	'
	'==============================================================================================
	Private Function K_Delay(oRow)
		Dim iRetVal, sDuration
		' Get parameters from spreadsheet as strings
		sDuration = CStr(oRow.Cells(1, XL_PARM_001).Value)
		' Check for numeric Duration
		If IsNumeric(sDuration) Then
			' WaitTime for sDuration seconds
			WaitTime CInt(sDuration)

			iRetVal = XL_DISPATCH_PASS
		Else

			' Duration is not numeric
			iRetVal = XL_DISPATCH_FAIL
		End If

		' Set return value
		K_Delay = iRetVal
	End Function

	'==============================================================================================
	' Function:		K_Exec
	' Purpose:		Load and execute another workbook
	'
	' Parameters:	Parm001 - Workbook - workbook to execute
	'
	' Returns:		TRUE  - if exec successful
	'				FALSE - if exec fails
	'
	'==============================================================================================
	Public Function K_Exec(oRow)
		Dim iRetVal, sWorkbook,aTempBuffer
		Dim sCodePath,i
		Dim oDialog, bDialogResult, iYesNo

		iYesNo = vbYes

		' Get parameters from spreadsheet as strings
		sWorkbook = CStr(oRow.Cells(1, XL_PARM_001).Value)
		' if the full path of the workbook has not been included
		' then work off the default directory
		'Test script directory
		if Right(sWorkbook, 3) = "xls"	then	'xls supplied
			'find out if full path has been passed from the spreadsheet
			aTempBuffer=split (sWorkbook,":\",-1)
			if ubound(aTempBuffer) > 0 then
				'Full path exists
				'So do nothing to path and directory
			else
				sWorkbook = sVBSFrameworkDir & "\" & sWorkbook
			end if
		else
			bDialogResult = False
			Set oDialog = CreateObject("UserAccounts.CommonDialog")
			do while bDialogResult = False
				oDialog.Filter =cstr(".xls")
				aTempBuffer=split (sWorkbook,":\",-1)
				if ubound(aTempBuffer) >0 then
					oDialog.InitialDir = CStr(oRow.Cells(1, XL_PARM_001).Value)
				else
					oDialog.InitialDir = sVBSFrameworkDir & "\" & sWorkbook
				end if 
				bDialogResult = oDialog.ShowOpen
				if bDialogResult = False then
					iYesNo = msgbox("You did not select a file." & Chr(13) & "Press Yes to try again or No to stop the test.", vbYesNo)
					if iYesNo = vbNo then
						exit do
					end if
				end if
		    	loop

			sWorkbook = oDialog.FileName
			set oDialog = Nothing
		    oRow.Cells(1, XL_LOG).Value = sWorkbook
		end if

		' Save current workbook just in case
		iRetVal = SaveWorkbook(oRow.Parent.Parent)

		if iYesNo = vbyes then
			' Recurse into a new instance of Main using sWorkbook
			Call Main(sWorkbook)
			' Assume all OK
			iRetVal = XL_DISPATCH_PASS
		else
			iRetVal = XL_DISPATCH_FAIL
		end if

		' Set return value
		K_Exec = iRetVal
	End Function

	'==============================================================================================
	' Function:		K_Save
	' Purpose:		Save the current workbook
	'
	' Parameters:	<none>
	'
	' Returns:		TRUE  - if save successful
	'				FALSE - if save fails
	'
	'==============================================================================================
	Private Function K_Save(oRow)
		Dim iRetVal

		' Save workbook
		iRetVal = SaveWorkbook(oRow.Parent.Parent)

		' Set return value
		K_Save = iRetVal
	End Function
	'==============================================================================================
	' Function:		K_LogLevel
	' Purpose:		Enable/Disable desktop alerts
	'
	' Parameters:	Parm001 - True or False
	'
	' Returns:		XL_DISPATCH_PASS  - if successful
	'
	'
	'==============================================================================================
	Private Function K_LogLevel(oRow)
		Dim iRetVal

		bLogLevel=true
		If (oFso.FileExists(sVBSFrameworkDir&"\"&sLogFileName)) Then
			oFso.DeleteFile(sVBSFrameworkDir&"\"&sLogFileName)
		end if
		iRetVal = XL_DISPATCH_PASS
		' Set return value
		K_LogLevel = iRetVal
	End Function

	'==============================================================================================
	' Function:		LoadVBS
	' Purpose:		Load and create instance of APP specific classes
	'
	' Parameters:	PARM_001 - APPCode - code indicating APP class to load
	'
	'==============================================================================================
	Private Function LoadVBS(oRow)
		Dim iRetVal, sAPPCode,sDescObj,sRepository,aTempBuffer
		Dim sCodePath,i,oTestObject, sAppFile, sAppName
		Dim oQTP
		
		Const sRoutine = "clsVBSFramework.LoadVBS"
		oTraceLog.Entered(sRoutine)

		sAppFile = CStr(oRow.Cells(1, XL_PARM_001).Value)
		sAPPCode = oFSO.getbasename(sAppFile)

		If oTestObjects.AddFile(UCase(sAppFile)) = False Then	'False means already loaded
		'If oTestObjects.AddFile(sAPPCode) = False Then	'False means already loaded
			' File already loaded
			'iRetVal = XL_DISPATCH_FAIL
			iRetVal = XL_DISPATCH_SKIP
			oRow.Cells(1, XL_LOG) = "Already loaded. Skipped."
		Else
			'Load APP codes (Functions,description objects,shared obj repository etc)
			
			aTempBuffer=split(sVBSFrameworkDir,"\",-1)
			For i=0 to ubound(aTempBuffer)-1
				sCodePath=sCodePath & aTempBuffer(i)&"\"
			Next
			' Get parameters from spreadsheet as strings
			
			sDescObj= CStr(oRow.Cells(1, XL_PARM_002).Value)
			sRepository=CStr(oRow.Cells(1, XL_PARM_003).Value)
			'handle specific code

			'APP keyword functions
			if sAPPCode="" then

			else
				'find out if full path has been passed from the spresdsheet
				aTempBuffer=split (sAppFile,":\",-1)
				if ubound(aTempBuffer) >0 then
					'Full path exists
					ExecuteFileToGlobal sAppFile
				else
					ExecuteFileToGlobal sVBSFrameworkDir & "\" & sAppFile
				end if
			end if

			'App Description objects
			if sDescObj="" Then 
				'description objects not required
			else
				'find out if full path has been passed from the spresdsheet
				aTempBuffer=split (sDescObj,":\",-1)
				if ubound(aTempBuffer) >0 then
					'Full path exists
					ExecuteFileToGlobal sDescObj
				else
					'Assume that only the decription objects file name was passed
					ExecuteFileToGlobal sCodePath & sDescObj
				end if
			end if


			'REM changed to have local oQTP object
			'App Repository objects
			if sRepository="" then
				'Repositorydescription objects not required
			else
				'find out if full path has been passed from the spreadsheet
				aTempBuffer=split (sRepository,":\",-1)
				Set oQTP = CreateObject("QuickTest.Application")
				if ubound(aTempBuffer) > 0 then
					'Full path exists
					 'Load the .tsr shared object repository
					oQTP.Test.Settings.Resources.ObjectRepositoryPath = sRepository
				else
					'Assume that only the decription objects file name was passed
					oQTP.Test.Settings.Resources.ObjectRepositoryPath = sCodePath & sRepository
				end If
				Set oQTP = Nothing
			end if

			iRetVal = XL_DISPATCH_PASS
		End If

		LoadVBS = iRetVal

		oTraceLog.Exited(sRoutine)
	End Function

	'==============================================================================================
	' Function:		LoadModuleAction
	'==============================================================================================
	Private Function LoadModuleAction(oRow)
		Dim iRetVal
		Dim sFilename

		Const sRoutine = "clsVBSFramework.LoadModuleAction"
		oTraceLog.Entered(sRoutine)
		iRetVal = XL_DISPATCH_PASS

		sFilename = CStr(oRow.Cells(1, XL_PARM_001).Value)
		
		If LoadModule(sFilename) = False Then iRetVal = XL_DISPATCH_FAIL
			
		LoadModuleAction = iRetVal
		oTraceLog.Exited(sRoutine)
	End Function

	'==============================================================================================
	' Function:		LoadModule
	'==============================================================================================
	Public Function LoadModule(sFilename)
		Dim iRetVal, bRetVal
		Dim sFullFilename

		Const sRoutine = "clsVBSFramework.LoadModule"
		oTraceLog.Entered(sRoutine)
		iRetVal = XL_DISPATCH_PASS
		bRetVal = True
	
		'Set oModules = New clsModules
		'Set oModules = Nothing
		'REM added Log
		oTraceLog.Message "Loading module: " & sFilename, LOG_MESSAGE
		
		If oTestObjects.AddFile(UCase(sFilename)) = False Then	'For compatibility with LoadVBS
			'File already loaded
			iRetVal = XL_DISPATCH_FAIL
			bRetVal = False
		
		Else
		
			'Add the source file name to the Modules list
			If oModules.GetSourceFile(sFilename) = "" Then
				oModules.SetSourceFile sFilename, sFilename
			End If

			'Check that the module file actually exists.
			'Get the full filename and path either a "\\" or a "D:\" indicates full filename and path
			
			If Left(Trim(sFilename),2) = "\\" OR Mid(Trim(sFilename),2,1) = ":" Then
				
				sFullFilename = IncludeFiles(sFilename, sVBSFrameworkDir)
			Else
				sFullFilename = IncludeFiles(sVBSFrameworkDir & "\" & sFilename, sVBSFrameworkDir)
			End If 
		
			'MsgBox sFullFilename		
			If oFso.FileExists(sFullFilename) Then
				oModules.SetLoaded sFilename
				'MsgBox sFullFilename
				bRetVal = ExecuteFileToGlobal(sFullFilename)
				If bRetVal Then
					oTraceLog.Message "Module loaded", LOG_MESSAGE 
					iRetVal = XL_DISPATCH_PASS
				Else
					oTraceLog.Message "Loading failed", LOG_ERROR 
					iRetVal = XL_DISPATCH_FAIL
				End If	
					
			Else
				'File does not exit
				iRetVal = XL_DISPATCH_FAIL
				bRetVal = False
				oTraceLog.Message " Loading failed. File " & sFullFilename & " does not exist", LOG_ERROR 
			End If

'			bRetVal = True
'			iRetVal = XL_DISPATCH_PASS
		
		End If

		LoadModule = bRetVal
	
		oTraceLog.Exited(sRoutine)
	End Function
	
	'==============================================================================================
	' Function:	DesignModeUpdateAttachments
	' Purpose:	
	' Parameters:	
	'==============================================================================================
	Public Function DesignModeUpdateAttachments(sFilesToBeRemoved, sFilesToBeAdded, sMaxFileNameLegth, sOutput)
		Dim iRetVal, iCounter, iMaxFileNameLegth
		Dim arrFilesToBeRemoved, arrFilesToBeAdded, arrOldNewFileName, arrPathAndNewFileName
		Dim oTest, oAttachTestFact, oAttachCurrentRunFact, oAttachList, oAttachment, oRegExp, oNewAttachment, sAttachmentOriginalName
		Dim fVersionControl
		Dim sAttachmentFileName, sAttachmentFullFileName, sLoadPath, sSuffix, sFileName, sNewFileName
		Dim bAttachmentExist
		
		Const sRoutine = "clsVBSFramework.DesignModeUpdateAttachments"
		oTraceLog.Entered(sRoutine)
		iRetVal = XL_DISPATCH_PASS
		
		If sMaxFileNameLegth="" Then
			iMaxFileNameLegth = 64
		Else
			iMaxFileNameLegth = CInt(sMaxFileNameLegth)
		End If
		
		Set oTest = globCurrentTSTest.Test
		Set oAttachTestFact = oTest.Attachments
		Set oAttachCurrentRunFact = globCurrentRun.Attachments
		
		Set fVersionControl = oTest.VCS
		If Not (sFilesToBeRemoved = "" And sFilesToBeAdded = "") Then
			If oTest.Field("TS_VC_STATUS")= "Checked_In" Then
				fVersionControl.CheckOut -1, "", False
			Else
				If UCase(oTest.Field("TS_VC_USER_NAME"))= UCase(globTDConnection.UserName) Then
					sOutput = sOutput & "I: Test was already checked out, no need for check in." & vblf
				Else
					sOutput = sOutput & "E: Test is already checked out by another user: "& oTest.Field("TS_VC_USER_NAME") & vblf
					DesignModeUpdateAttachments = XL_DISPATCH_FAIL : oTraceLog.Exited(sRoutine) : Exit Function
				End If
			End If
		End If
		
		'To Delete Attachments
		If sFilesToBeRemoved = "" Then
			sOutput = sOutput & "D: Nothing to be deleted." & vblf
		Else
			arrFilesToBeRemoved = Split(sFilesToBeRemoved, ",")
			For iCounter = 0 To UBound(arrFilesToBeRemoved)
				Set oRegExp = New RegExp
				oRegExp.Pattern = "^" & arrFilesToBeRemoved(iCounter)
				oRegExp.IgnoreCase = False
				Set oAttachList = oAttachTestFact.NewList("")
				bAttachmentExist = false
				If oAttachList.Count > 0 Then
					For Each oAttachment In oAttachList
						sAttachmentFullFileName = Right(oAttachment.FileName, len(oAttachment.FileName) - InstrRev(oAttachment.FileName,"\"))
						sAttachmentFileName = Right(sAttachmentFullFileName, len(sAttachmentFullFileName) - Instr(sAttachmentFullFileName,"_"))
						sAttachmentFileName = Right(sAttachmentFileName, len(sAttachmentFileName) - Instr(sAttachmentFileName,"_"))
						If oRegExp.Test(sAttachmentFileName) Then
							oAttachTestFact.RemoveItem(oAttachment.ID)
							bAttachmentExist = True
							sOutput = sOutput & "D: DELETED Attachment '" & sAttachmentFileName & "'." & vblf
						End If
					Next
					Set oAttachment = Nothing
				End If
				Set oAttachList = Nothing
				If Not bAttachmentExist Then
					sOutput = sOutput & "W: No Attachment for condition: '" & arrFilesToBeRemoved(iCounter) & "'." & vblf
				End If
			Next
		End If
		
''--------------------OLD     To Upload Attachments--------------------
'		If sFilesToBeAdded = "" Then
'			sOutput = sOutput & "D: Nothing to be uploaded." & vblf
'		Else
'			arrFilesToBeAdded = Split(sFilesToBeAdded, "|")
'			Set oAttachList = oAttachCurrentRunFact.NewList("")
'			For iCounter = 0 To UBound(arrFilesToBeAdded)
'				arrOldNewFileName = Split(arrFilesToBeAdded(iCounter),",")
'				Set oRegExp = New RegExp
'				oRegExp.Pattern = "^" & Replace(Replace(Replace(arrOldNewFileName(0),".","\."),"*",".*"),"?",".")
'				oRegExp.IgnoreCase = False
'				bAttachmentExist = False
'				For Each oAttachment In oAttachList
'					sAttachmentFullFileName = Right(oAttachment.FileName, len(oAttachment.FileName) - InstrRev(oAttachment.FileName,"\"))
'					sAttachmentFileName = Right(sAttachmentFullFileName, len(sAttachmentFullFileName) - Instr(sAttachmentFullFileName,"_"))
'					sAttachmentFileName = Right(sAttachmentFileName, len(sAttachmentFileName) - Instr(sAttachmentFileName,"_"))
					'MsgBox sAttachmentFileName & vblf & arrOldNewFileName(0) & vblf & oRegExp.Test(sAttachmentFileName)
'					sAttachmentOriginalName = sAttachmentFileName
'					If oRegExp.Test(sAttachmentFileName) Then
'						MsgBox oAttachment.FileName & vblf & oAttachment.Virtual
'						oAttachment.Load True, ""
'						oFso.CopyFile oAttachment.FileName, PATH_RESOURCES, True
'						sAttachmentFileName = Left(arrOldNewFileName(1),Instr(arrOldNewFileName(1),"(1)")-1) & sAttachmentFileName
'						sSuffix = Right(arrOldNewFileName(1), Len(arrOldNewFileName(1))-Instr(arrOldNewFileName(1),"(1)")-len("1)"))
'						If Instr(sSuffix,".")=0 Then
'							sAttachmentFileName = Left(sAttachmentFileName,Len(sAttachmentFileName)-Len(".xxx")) & sSuffix & Right(sAttachmentFileName,Len(".xxx"))
'						Else
'							sAttachmentFileName = Left(sAttachmentFileName,Len(sAttachmentFileName)-Len(".xxx")) & sSuffix
'						End If
'						If Len(sAttachmentFileName)>iMaxFileNameLegth Then
'							sAttachmentFileName = Left(Left(sAttachmentFileName,Len(sAttachmentFileName)-Len(".xxx")),iMaxFileNameLegth-Len(".xxx")) & Right(sAttachmentFileName,Len(".xxx"))
'						End If
'						If Not oFso.FileExists(PATH_RESOURCES & sAttachmentFileName) Then
'							oFSO.MoveFile PATH_RESOURCES & sAttachmentFullFileName, PATH_RESOURCES & sAttachmentFileName
'						End If
'						If Not oFso.FileExists(PATH_RESOURCES & sAttachmentFileName) Then
'							sOutput = sOutput & "E: File '" & sAttachmentFileName & "' not found in '"& PATH_RESOURCES &"'." & vblf
'							DesignModeUpdateAttachments = XL_DISPATCH_FAIL : oTraceLog.Exited(sRoutine) : Exit Function
'						End If
'						Set oNewAttachment = oAttachTestFact.AddItem(Null)
'						oNewAttachment.FileName = PATH_RESOURCES & sAttachmentFileName
'						oNewAttachment.Type = 1
'						oNewAttachment.Post
'						bAttachmentExist = true
'						sOutput = sOutput & "D: UPLOADED Attachment '" & sAttachmentFileName & "'." & vblf
'						Set oNewAttachment = Nothing 
'					End If
'				Next
'				Set oAttachment = Nothing
'				If Not bAttachmentExist Then
'					sOutput = sOutput & "W: No Attachment for condition: '" &  arrOldNewFileName(0) & "'." & vblf
'				End If
'			Next
'			Set oAttachList = Nothing
'		End If
''--------------------OLD     To Upload Attachments--------------------

		If sFilesToBeAdded = "" Then
			sOutput = sOutput & "D: Nothing to be uploaded." & vblf
		Else
			arrFilesToBeAdded = Split(sFilesToBeAdded, "|")
			For iCounter = 0 To UBound(arrFilesToBeAdded)
				arrPathAndNewFileName = Split(arrFilesToBeAdded(iCounter),",")
				If oFso.FileExists(arrPathAndNewFileName(0)) Then
					oFso.CopyFile arrPathAndNewFileName(0), PATH_RESOURCES, True
					sFileName = Right(arrPathAndNewFileName(0), len(arrPathAndNewFileName(0)) - InstrRev(arrPathAndNewFileName(0),"\"))
					sNewFileName = Left(arrPathAndNewFileName(1),Instr(arrPathAndNewFileName(1),"(1)")-1) & sFileName
					sSuffix = Right(arrPathAndNewFileName(1), Len(arrPathAndNewFileName(1))-Instr(arrPathAndNewFileName(1),"(1)")-len("1)"))
					If Instr(sSuffix,".")=0 Then
						sNewFileName = Left(sNewFileName,Len(sNewFileName)-Len(".xxx")) & sSuffix & Right(sNewFileName,Len(".xxx"))
					Else
						sNewFileName = Left(sNewFileName,Len(sNewFileName)-Len(".xxx")) & sSuffix
					End If
					If Len(sNewFileName)>iMaxFileNameLegth Then
						sNewFileName = Left(Left(sNewFileName,Len(sNewFileName)-Len(".xxx")),iMaxFileNameLegth-Len(".xxx")) & Right(sNewFileName,Len(".xxx"))
					End If
					If oFso.FileExists(PATH_RESOURCES & sNewFileName) Then
						oFso.DeleteFile PATH_RESOURCES & sNewFileName
						oFso.MoveFile PATH_RESOURCES & sFileName, PATH_RESOURCES & sNewFileName
					Else
						oFso.MoveFile PATH_RESOURCES & sFileName, PATH_RESOURCES & sNewFileName
					End If
					Set oNewAttachment = oAttachTestFact.AddItem(Null)
					oNewAttachment.FileName = PATH_RESOURCES & sNewFileName
					oNewAttachment.Type = 1
					oNewAttachment.Post
					sOutput = sOutput & "D: UPLOADED Attachment '" & sNewFileName & "'." & vblf
					oFso.DeleteFile PATH_RESOURCES & sNewFileName
					Set oNewAttachment = Nothing 
				Else
					sOutput = sOutput & "W: File '" & arrPathAndNewFileName(0) & "' not found." & vblf
				End If
			Next
		End If
				
		'If Not (sFilesToBeRemoved = "" And sFilesToBeAdded = "") Then
		'	oTest.VCS.CheckIn "", ""
		'End If
		
		Set oAttachTestFact = Nothing
		Set oTest = Nothing
		oTraceLog.Message(Array("sOutput", LOG_SLOG))
		DesignModeUpdateAttachments = iRetVal
		oTraceLog.Exited(sRoutine)
	End Function
	
	'==============================================================================================
	' Function:	DesignModeUpdateAttachmentsAction
	' Purpose:	
	' Parameters:	
	'==============================================================================================
	Public Function DesignModeUpdateAttachmentsAction(oRow, sOutput)
		Dim iRetVal
		Dim sFilesToBeRemoved, sFilesToBeAdded, sMaxFileNameLegth
		'Dim oQC
		
		Const sRoutine = "clsVBSFramework.DesignModeUpdateAttachmentsAction"
		oTraceLog.Entered(sRoutine)
		iRetVal = XL_DISPATCH_PASS
		
		'check if DESIGN mode is on
		'Set oQC = New clsQC
  		If oQC.IsQCRun() Then 
   			If Not globCurrentTSTest.Test.Field("TS_STATUS") = "Design" Then
   				sOutput = sOutput & "D: Not in 'Design' mode." & vblf
   				DesignModeUpdateAttachmentsAction = XL_DISPATCH_SKIP : oTraceLog.Exited(sRoutine) : Exit Function
   			End If
   		Else
   			sOutput = sOutput & "D: Test is not running from Quality Center." & vblf
   			DesignModeUpdateAttachmentsAction = XL_DISPATCH_SKIP : oTraceLog.Exited(sRoutine) : Exit Function
   		End If
		
		sFilesToBeRemoved = ""
		sFilesToBeAdded = ""
		
		sFilesToBeRemoved = Trim(CStr(oRow.Cells(1, XL_PARM_001).Value))
		sFilesToBeRemoved = Replace(Replace(Replace(sFilesToBeRemoved,".","\."),"*",".*"),"?",".")
		sFilesToBeAdded = Trim(CStr(oRow.Cells(1, XL_PARM_002).Value))
		sMaxFileNameLegth = Trim(CStr(oRow.Cells(1, XL_PARM_003).Value))
		
		DesignModeUpdateAttachmentsAction = DesignModeUpdateAttachments(sFilesToBeRemoved, sFilesToBeAdded, sMaxFileNameLegth, sOutput)
		oTraceLog.Exited(sRoutine)
	End Function

	'==============================================================================================
	' Function:	DesignModeUpdateAttachmentsAction
	' Purpose:	
	' Parameters:	
	'==============================================================================================
	Public Function DesignModeEnd(oRow, sOutput)
		Dim iRetVal
		'Dim oQC
		
		Const sRoutine = "clsVBSFramework.DesignModeEnd"
		oTraceLog.Entered(sRoutine)
		iRetVal = XL_DISPATCH_PASS
		
		'TODO: check if DESIGN mode is on
		'Set oQC = New clsQC
  		If oQC.IsQCRun() Then 
   			If Not globCurrentTSTest.Test.Field("TS_STATUS") = "Design" Then
   				sOutput = sOutput & "D: Not in 'Design' mode." & vblf
   				DesignModeEnd = XL_DISPATCH_SKIP : oTraceLog.Exited(sRoutine) : Exit Function
   			End If
   		Else
   			sOutput = sOutput & "D: Test is not running from Quality Center." & vblf
   			DesignModeEnd = XL_DISPATCH_SKIP : oTraceLog.Exited(sRoutine) : Exit Function
   		End If
		
		iRetVal=XL_DISPATCH_END
		DesignModeEnd = iRetVal
		oTraceLog.Exited(sRoutine)
	End Function
	
	'==============================================================================================
	' Function:	ExecuteStatementAction
	' Purpose:	
	' Parameters:	
	'==============================================================================================
	Public Function ExecuteStatement(oRow, sLog)
		Dim iRetVal
		Dim sStatement

		Const sRoutine = "clsVBSFramework.ExecuteStatement"

		oTraceLog.Entered(sRoutine)
		iRetVal = XL_DISPATCH_PASS

		sStatement = Trim(CStr(oRow.Cells(1, XL_PARM_001).Value))
		'MsgBox sStatement
		
		Execute sStatement
		
		'MsgBox "sLog in ExecuteStatement = " & sLog 
		
		ExecuteStatement = iRetVal
		oTraceLog.Exited(sRoutine)
	End Function

	'==============================================================================================
	' Function:	k_selectdatatable
	' Purpose:	Pauses script execution to choose the appropriate datatable.
	'
	' Parameters:	PARM_001 - APPCode - code indicating APP class to load
	'
	'==============================================================================================
	Private function k_SelectDataTable(oRow)
		Dim objDialog
		Dim defaultFile

	    Set objDialog = CreateObject("UserAccounts.CommonDialog")

	   	 With objDialog
	       	.Filter =cstr(".xls")
	'       .Flags = "&H0200"
	       	.InitialDir = CStr(oRow.Cells(1, XL_PARM_001).Value)
		    .FileName = defaultFile
		    .ShowOpen
	    End With

	    oRow.Cells(1, XL_PARM_001).Value = objDialog.FileName

	    'Msgbox objDialog.FileName

	    K_Exec(oRow)

	    Set objDialog = Nothing
	End Function
	
	'==============================================================================================
	' Function:		K_SelectItem
	' Purpose:		uses the Open Dialog to populate the output field from a comma seperated list
	'
	' Parameters:	Parm001 - Is a comma seperated list values
	'
	' Returns:		TRUE  - if exec successful
	'				FALSE - if exec fails
	'
	'==============================================================================================
	Public Function K_SelectItem(oRow)
		Dim iRetVal, sItemList,arrItems, sItem
		Dim fso, sTmpFolder, oTmpFolder, oItem
		Dim sCodePath,i
		Dim oDialog, bDialogResult, iYesNo
		
		iYesNo = vbYes
		sTmpFolder = "h:\tmpselectitem"
		Set fso = CreateObject("Scripting.FileSystemObject")
		
		'Create folder
		
		If (fso.FolderExists(sTmpFolder)) Then
			fso.DeleteFolder(sTmpFolder)
		End If

		fso.CreateFolder(sTmpFolder)
		
		' Get parameters from spreadsheet as strings
		sItemList = CStr(oRow.Cells(1, XL_PARM_001).Value)
		
		arrItems = Split(sItemList,",",-1,1)

		For Each sItem in arrItems
		   fso.OpenTextFile sTmpFolder & "\" & sItem, 2, True		
		Next
		
		'Select file using open dialog

		bDialogResult = False
		Set oDialog = CreateObject("UserAccounts.CommonDialog")
		do while bDialogResult = False
			oDialog.InitialDir = sTmpFolder
			bDialogResult = oDialog.ShowOpen
			if bDialogResult = False then
				iYesNo = msgbox("You did not select a file." & Chr(13) & "Press Yes to try again or No to stop the test.", vbYesNo)
				if iYesNo = vbNo then
					exit do
				end if
			end if
		loop

		sItem = UCase(fso.GetFileName(oDialog.FileName))
		Set fso = Nothing
		set oDialog = Nothing
		oRow.Cells(1, XL_LOG).Value = sItem

		if iYesNo = vbyes then
			iRetVal = XL_DISPATCH_PASS
		else
			iRetVal = XL_DISPATCH_FAIL
		end if

		' Set return value
		K_SelectItem = iRetVal
	End Function

	'==============================================================================
	' Function:		
	' Purpose:
	'==============================================================================
	'REM QTP specific
'	Public Function GetObjectSyncTimeOut()
'		GetObjectSyncTimeOut = oQTP.Test.Settings.Run.ObjectSyncTimeOut 
'	End Function
	
'	Public Sub SetObjectSyncTimeOut(lMilliSeconds)
'		oQTP.Test.Settings.Run.ObjectSyncTimeOut = lMilliSeconds
'	End Sub
	
	'==============================================================================================
	' Function/Sub:
	'==============================================================================================
	Public Sub WriteIntoLogAndOutputParamsCells(oRow, sLog, sOutputParams)
		If sLog <> "" Then
			If oRow.Cells(1, XL_LOG).Value = "" Then
				oRow.Cells(1, XL_LOG).Value = sLog
			ElseIf Right(oRow.Cells(1, XL_LOG).Value, 1) = vblf Then
				oRow.Cells(1, XL_LOG).Value = oRow.Cells(1, XL_LOG).Value & sLog '& vblf & sLog
			Else
				oRow.Cells(1, XL_LOG).Value = oRow.Cells(1, XL_LOG).Value & vblf & sLog
			End If
		End If
		If sOutputParams <> "" Then
			If oRow.Cells(1, XL_OUTPUT_PARAMS).Value = "" Then
				oRow.Cells(1, XL_OUTPUT_PARAMS).Value = sOutputParams
			ElseIf Right(oRow.Cells(1, XL_OUTPUT_PARAMS).Value, 1) = "," Then
				oRow.Cells(1, XL_OUTPUT_PARAMS).Value = oRow.Cells(1, XL_OUTPUT_PARAMS).Value & sOutputParams
			Else
				oRow.Cells(1, XL_OUTPUT_PARAMS).Value = oRow.Cells(1, XL_OUTPUT_PARAMS).Value & "," & sOutputParams
			End If
		End If
	End Sub
	
	'==============================================================================================
	' Function/Sub:
	'==============================================================================================
'	Sub PrecacheExcelToDictionary(sResourceName, sLog)
	Sub PrecacheExcelToDictionary(sResourceName)
		Dim oExcel, oWB, oSheet
		Dim dictSelectedExternalData, dictMapping
		Dim sLocalFile, sMappingFolderNameUC, sMappingSheetUC, sSheetNameUC
		Dim iRows, iColumns, iRow, iColumn, iCounter
		Dim arrMappingNewSheets
		ReDim arrContentRC(-1,-1)
		Const sRoutine = "clsVBSFramework.PrecacheExcelToDictionary"
		oTraceLog.Entered(sRoutine)

		'todo change path to globalDataFOlder
		'STOP
		'sLocalFile = PATH_TESTS & sResourceName & ".xls"
		sLocalFile = globDataPath & "\" & sResourceName & ".xls"

'msgbox globDataPath

		If globDebug Then
			oTraceLog.Message "Local file set to: " & sLocalFile, LOG_DEBUG
		End If
		
		If oFso.FileExists(sLocalFile) Then
			'sLog = sLog & "D: Precaching data from local external file: " & sLocalFile & vbLf
			oTraceLog.Message "D: Precaching data from local external file: " & sLocalFile, LOG_MESSAGE				
		ElseIf oQC.IsQCRun Then
			'Set oQC = New clsQC
			sLocalFile = oQC.GetResourceFileFromQC(sResourceName)
			If sLocalFile = "" Then
				'sLog = sLog & "E: Resource file '" & sResourceName & "' not found neither locally nor in HPQC Resources. Check file name, file path or HPQC connection." & vblf
				oTraceLog.Message "Resource file '" & sResourceName & "' not found neither locally nor in HPQC Resources. Check file name, file path or HPQC connection.", LOG_ERROR
				oTraceLog.Exited(sRoutine)
				'Set oQC = Nothing
				
				Exit Sub
			End If
		'	sLog = sLog & "D: Precaching data from QC Resources file: " & sLocalFile & vbLf
			oTraceLog.Message "D: Precaching data from QC Resources file: " & sLocalFile, LOG_MESSAGE
			'Set oQC = Nothing
		Else
			'TODO: proper error message
			'sLog = sLog & "E: Resource file '" & sResourceName & "' not found neither locally nor in HPQC Resources. Check file name, file path or HPQC connection." & vbLf
			oTraceLog.Message "Resource file '" & sResourceName & "' not found neither locally nor in HPQC Resources. Check file name, file path or HPQC connection.",LOG_ERROR
			oTraceLog.Exited(sRoutine)
			Exit Sub
		End If
		
		Set dictSelectedExternalData = CreateObject("scripting.dictionary")
		Set dictMapping = CreateObject("scripting.dictionary")
		
		Set oExcel = CreateObject("Excel.Application")
		oExcel.DisplayAlerts = False
		oExcel.Visible = False
		
		Set oWB = oExcel.Workbooks.Open(sLocalFile)

		'process MAPPING sheet if exists
		For Each oSheet In oWB.Worksheets
			If UCase(oSheet.Name) = "MAPPING" Then
				'determine real used range
				On Error Resume Next
				iRows = oSheet.Cells.Find("*", oSheet.Range("A1"), -4163, 2, 1, 2).Row
				If Err.Number <> 0 Then
					iRows = 0
					Err.Clear
				End If
				On Error Goto 0

				For iRow = 1 To iRows
					'column 1 - QC folder name
					'column 2 - sheet to use
					sMappingFolderNameUC = UCase(oSheet.Cells(iRow, 1).Value)
					sMappingSheetUC = UCase(oSheet.Cells(iRow, 2).Value)
					
					If sMappingFolderNameUC <> "" And sMappingSheetUC <> "" Then
						If Not dictMapping.Exists(sMappingSheetUC) Then
							dictMapping.Add sMappingSheetUC, sMappingFolderNameUC
						Else
							dictMapping(sMappingSheetUC) = dictMapping(sMappingSheetUC) & "," & sMappingFolderNameUC
						End If
					End If
				Next
			End If
		Next
		
		For Each oSheet In oWB.Worksheets
			'determine real used range
			On Error Resume Next
			iRows = oSheet.Cells.Find("*", oSheet.Range("A1"), -4163, 2, 1, 2).Row
			If Err.Number <> 0 Then
				iRows = 0
				Err.Clear
			End If
			iColumns = oSheet.Cells.Find("*", oSheet.Range("A1"), -4163, 2, 2, 2).Column
			If Err.Number <> 0 Then
				iColumns = 0
				Err.Clear
			End If
			On Error Goto 0
			'MsgBox "Rows: " & iRows & ", Columns: " & iColumns
			
			ReDim arrContentRC(iRows - 1, iColumns - 1)
			
			'fill array with data from spreadsheet
			For iRow = 1 To iRows
				For iColumn = 1 To iColumns
					'text limitation 1024 chars
					If Len(oSheet.Cells(iRow, iColumn)) > 1020 Then
						arrContentRC(iRow - 1, iColumn - 1) = oSheet.Cells(iRow, iColumn).Value
					Else
						arrContentRC(iRow - 1, iColumn - 1) = oSheet.Cells(iRow, iColumn).Text
					End If
				Next
			Next
			
			sSheetNameUC = UCase(oSheet.Name)
			
			'add sheet into dictionary if not already there
			If Not dictSelectedExternalData.Exists(sSheetNameUC) Then
				dictSelectedExternalData.Add sSheetNameUC, arrContentRC
			End If
			
			'add new sheets if mapping exist
			'TODO: if entry exists in mapping and also as separate sheet there's no guarantee which one has the preference (depends on Excel sheet order mechanism)
			If dictMapping.Exists(sSheetNameUC) Then
				arrMappingNewSheets = Split(dictMapping(sSheetNameUC), ",")
				For iCounter = 0 To UBound(arrMappingNewSheets)
					If Not dictSelectedExternalData.Exists(arrMappingNewSheets(iCounter)) Then
						dictSelectedExternalData.Add arrMappingNewSheets(iCounter), arrContentRC
					End If
				Next
			End If
		Next
		
		dictExternalData.Add UCase(sResourceName), dictSelectedExternalData
		
		Set dictSelectedExternalData = Nothing
		Set dictMapping = Nothing
		oWB.Close
		oExcel.Quit
		Set oExcel = Nothing
		oTraceLog.Exited(sRoutine)
	End Sub
	
	'==============================================================================================
	' Function/Sub:
	'==============================================================================================
	Public Function ResolveExternalData(ByVal sResourceName, ByVal sLookupColumn, ByVal sLookupValue, ByVal sResultColumn, sExecutionPath, sLog)
		Dim sResult, sExecPathFolder
		Dim iExecPathCounter, iRows, iColumns, iRow, iLookupColumn, iResultRow, iResultColumn, iHeaderColumn
		Dim dictSelectedExternalData
		Dim arrExecutionPath, arrSheet, arrSheets

		sResourceName = UCase(Trim(sResourceName))
		sLookupColumn = UCase(Trim(sLookupColumn))
		sResultColumn = UCase(Trim(sResultColumn))
		sLookupValue = UCase(Trim(sLookupValue))
		sResult = "%NOT FOUND%"
		
		Set dictSelectedExternalData = dictExternalData(sResourceName)
		
		If UBound(dictSelectedExternalData.Keys) = 0 Then
			'resolve values from first sheet if it's the only one
			arrSheets = dictSelectedExternalData.Keys
			arrExecutionPath = Split(UCase(arrSheets(0) & "\" & "Sheet1\Mappe1\GLOBAL\Default\" & sExecutionPath), "\")
			Erase arrSheets 'clears the array
		Else
			arrExecutionPath = Split(UCase("Sheet1\Mappe1\GLOBAL\Default\" & sExecutionPath), "\")
		End If
	
		'loop from the innermost folder to the root
		For iExecPathCounter = UBound(arrExecutionPath) To 0 Step -1
			
			sExecPathFolder = arrExecutionPath(iExecPathCounter)
			
			If dictSelectedExternalData.Exists(sExecPathFolder) Then
				arrSheet = dictSelectedExternalData(sExecPathFolder)
				iRows = UBound(arrSheet, 1)
				iColumns = UBound(arrSheet, 2)
	
				iLookupColumn = -1
				iResultRow = -1
				iResultColumn = -1
				
				If IsNumeric(sLookupColumn) Then
					iLookupColumn = CInt(sLookupColumn) - 1
				Else
					For iHeaderColumn = 0 To iColumns
						If UCase(arrSheet(0, iHeaderColumn)) = sLookupColumn Then
							iLookupColumn = iHeaderColumn
							Exit For
						End If
					Next
				End If
	
				If iLookupColumn > -1 Then
					For iRow = 1 To iRows
						If UCase(arrSheet(iRow, iLookupColumn)) = sLookupValue Then
							iResultRow = iRow
							Exit For
						End If
					Next
	
					If iResultRow > -1 Then
					
						If IsNumeric(sResultColumn) Then
							iResultColumn = CInt(sResultColumn) - 1
						Else
							For iHeaderColumn = 0 To iColumns
								If UCase(arrSheet(0, iHeaderColumn)) = sResultColumn Then
									iResultColumn = iHeaderColumn
									Exit For
								End If
							Next
						End If
						
						If iResultColumn > -1 Then
	                		sResult = arrSheet(iResultRow, iResultColumn)	                		
	                		sLog = sLog & "D: Value resolved from sheet: " & sExecPathFolder & vblf
						End If
					End if
				End If
			End If
				
			'exit loop if value already resolved
			If sResult <> "%NOT FOUND%" Then Exit for
		
		Next
	
		ResolveExternalData = sResult
	End Function

	'==============================================================================================
	' Function/Sub:
	'==============================================================================================	
	Public Function ValidateWorkbook(oWorkbook)
		Dim arrColDef(17, 3), arrErrors(), iErrorCount
		Dim iRow, iColumn, iRowCount, iColumnCount, iCounter
		Dim oFS, oSheet, oCell, oRegEx, oMatches, oMatch, oNames
		Dim dictNames, dictGlobal, dictLocal, dictReferences, dictMapping
		Dim arrTSParentFolders, sTSParentFolder
		Dim arrStreams, sStream
		Dim sValueFoundSection, sValueFound, sParameterName, sReplacePattern
		Dim dictUniqueConfigParams, arrUniqueConfigParamsKeys
		Dim bLongCellPresent
		Dim sError, sErrorReplaced, sKeyword, sCellValue, sFirstAddress, sLog
		
		Const sRoutine = "clsVBSFramework.ValidateWorkbook"
		oTraceLog.Entered(sRoutine)
		
		iErrorCount = 0
		
		'set the mandatory columns
		arrColDef(0, 0) = "XL_DISABLE"			:arrColDef(0, 1) = "M"		:arrColDef(0, 3) = "disable"
		arrColDef(1, 0) = "XL_RESULT"			:arrColDef(1, 1) = "M"		:arrColDef(1, 3) = "result"
		arrColDef(17, 0) = "XL_STREAM"			:arrColDef(17, 1) = "O"		:arrColDef(17, 3) = "stream"
		arrColDef(2, 0) = "XL_KEYWORD"			:arrColDef(2, 1) = "M"		:arrColDef(2, 3) = "keyword"
		arrColDef(3, 0) = "XL_COMMENT"			:arrColDef(3, 1) = "M"		:arrColDef(3, 3) = "comment"
		arrColDef(4, 0) = "XL_LOG"				:arrColDef(4, 1) = "M"		:arrColDef(4, 3) = "log"
		arrColDef(16, 0) = "XL_REFERENCE"		:arrColDef(16, 1) = "O"		:arrColDef(16, 3) = "reference"
		arrColDef(5, 0) = "XL_OUTPUT_PARAMS"	:arrColDef(5, 1) = "M"		:arrColDef(5, 3) = "output_params"
		arrColDef(6, 0) = "XL_PARM_001"			:arrColDef(6, 1) = "M"		:arrColDef(6, 3) = "parm_001"
		arrColDef(7, 0) = "XL_PARM_002"			:arrColDef(7, 1) = "M"		:arrColDef(7, 3) = "parm_002"
		arrColDef(8, 0) = "XL_PARM_003"			:arrColDef(8, 1) = "M"		:arrColDef(8, 3) = "parm_003"
		arrColDef(9, 0) = "XL_PARM_004"			:arrColDef(9, 1) = "M"		:arrColDef(9, 3) = "parm_004"
		arrColDef(10, 0) = "XL_PARM_005"		:arrColDef(10, 1) = "O"		:arrColDef(10, 3) = "parm_005"
		arrColDef(11, 0) = "XL_PARM_006"		:arrColDef(11, 1) = "O"		:arrColDef(11, 3) = "parm_006"
		arrColDef(12, 0) = "XL_PARM_007"		:arrColDef(12, 1) = "O"		:arrColDef(12, 3) = "parm_007"
		arrColDef(13, 0) = "XL_PARM_008"		:arrColDef(13, 1) = "O"		:arrColDef(13, 3) = "parm_008"
		arrColDef(14, 0) = "XL_PARM_009"		:arrColDef(14, 1) = "O"		:arrColDef(14, 3) = "parm_009"
		arrColDef(15, 0) = "XL_PARM_010"		:arrColDef(15, 1) = "O"		:arrColDef(15, 3) = "parm_0010"
		If globDebug Then
			oTraceLog.Message "arrColDef was set", LOG_DEBUG
		End If
		'TODO add check if all rows are using correctly, not using more columns as is defined in header
		If globDebug Then
			oTraceLog.Message "Count of Workeets: " & oWorkbook.Worksheets.Count, LOG_DEBUG
		End If
		For Each oSheet in oWorkbook.Worksheets
			iRowCount = oSheet.UsedRange.Rows.Count
			iColumnCount = oSheet.UsedRange.Columns.Count
			
			AssignHeaderColumnXLValues oSheet
			
			'check for #N/A or #NV values
			
			Set oCell = oSheet.UsedRange.Find("#N/A")
			
			If globDebug Then
				oTraceLog.Message "checking for #N/A", LOG_DEBUG
			End If
			
			If Not oCell Is Nothing Then
			    sFirstAddress = oCell.Address
			    Do
			    	If oSheet.Cells(oCell.Row, XL_DISABLE).Value = "" Then
				    	ReDim Preserve arrErrors(iErrorCount)
				    	arrErrors(iErrorCount) = "Sheet " & oSheet.Name & " [" & Chr(64 + oCell.Column) & "%ROW" & oCell.Row & "%]: #N/A value found!"
				    	iErrorCount = iErrorCount + 1
				    End If
			    	
			    	Set oCell = oSheet.UsedRange.FindNext(oCell)
			    	If Not oCell Is Nothing Then
						If oCell.Address = sFirstAddress Then Exit Do
					End If
			    Loop Until oCell Is Nothing
			    
			    Exit For
			End If
			
			Set oCell = oSheet.UsedRange.Find("#NV")
			If globDebug Then
				oTraceLog.Message "checking for #NV", LOG_DEBUG
			End If
			If Not oCell Is Nothing Then
			    sFirstAddress = oCell.Address
			    Do
			    	If oSheet.Cells(oCell.Row, XL_DISABLE).Value = "" Then
				    	ReDim Preserve arrErrors(iErrorCount)
				    	arrErrors(iErrorCount) = "Sheet " & oSheet.Name & " [" & Chr(64 + oCell.Column) & "%ROW" & oCell.Row & "%]: #NV value found!"
				    	iErrorCount = iErrorCount + 1
				    End If
			    	
			    	Set oCell = oSheet.UsedRange.FindNext(oCell)
			    	If Not oCell Is Nothing Then
						If oCell.Address = sFirstAddress Then Exit Do
					End If
			    Loop Until oCell Is Nothing
			    Exit For
			End If
			
			'loop thru all cells to find if there's long cell present (> 1020 chars)
			
			bLongCellPresent = False
			If globDebug Then
				oTraceLog.Message "checking for log cell", LOG_DEBUG
			End If
			For Each oCell In oSheet.UsedRange
				If Len(oCell) > 1020 Then bLongCellPresent = True
			Next
			
		
						
			
			'CHECK if all mandatory columns are present
		
			'reset column definitions to zero
			For iCounter = 0 To UBound(arrColDef, 1)
				arrColDef(iCounter, 2) = 0
			Next
			
			If globDebug Then
				oTraceLog.Message "Assigning columns id to column name ", LOG_DEBUG
			End If	
			For iCounter = 1 To iColumnCount
				sKeyword = oSheet.Cells(1, iCounter).Text
				Select Case UCase(sKeyword)
					Case "DISABLE", "SKIP"		arrColDef(0, 2) = iCounter
					Case "RESULT", "RES"		arrColDef(1, 2) = iCounter
					Case "STREAM", "STREAMS"	arrColDef(17, 2) = iCounter
					Case "KEYWORD", "KW"		arrColDef(2, 2) = iCounter
					Case "COMMENT"				arrColDef(3, 2) = iCounter
					Case "LOG"					arrColDef(4, 2) = iCounter
					Case "ID", "REFERENCE"		arrColDef(16, 2) = iCounter
					Case "OUTPUT_PARAMS"		arrColDef(5, 2) = iCounter
					Case "PARM_001"				arrColDef(6, 2) = iCounter
					Case "PARM_002"				arrColDef(7, 2) = iCounter
					Case "PARM_003"				arrColDef(8, 2) = iCounter
					Case "PARM_004"				arrColDef(9, 2) = iCounter
					Case "PARM_005"				arrColDef(10, 2) = iCounter
					Case "PARM_006"				arrColDef(11, 2) = iCounter
					Case "PARM_007"				arrColDef(12, 2) = iCounter
					Case "PARM_008"				arrColDef(13, 2) = iCounter
					Case "PARM_009"				arrColDef(14, 2) = iCounter
					Case "PARM_010"				arrColDef(15, 2) = iCounter
					
					'compatibility with older spreadsheets
					Case "OUTPUT"				arrColDef(4, 2) = iCounter
					Case "MESSAGE"				arrColDef(5, 2) = iCounter
				End Select
				
			Next
	
			If globDebug Then
				oTraceLog.Message "chceking for mandatory columns", LOG_DEBUG
			End If
			
			For iCounter = 0 To UBound(arrColDef, 1)
				If arrColDef(iCounter, 1) = "M" And arrColDef(iCounter, 2) = 0 Then
					ReDim Preserve arrErrors(iErrorCount)
					arrErrors(iErrorCount) = "Sheet " & oSheet.Name & ": Mandatory column missing in spreadsheet: " & arrColDef(iCounter, 3)
					iErrorCount = iErrorCount + 1
				End If
			Next
			
			'fill dictionary object with streams used in the spreadsheet
			If arrColDef(17, 2) <> 0 Then
				iColumn = arrColDef(17, 2)
				For Each oCell In oSheet.Range(oSheet.Cells(2, iColumn), oSheet.Cells(iRowCount, iColumn))
					'skip disabled rows
					If oSheet.Cells(oCell.Row, XL_DISABLE) = "" Then
						sCellValue = Trim(oCell.Text)
						If sCellValue <> "" Then
							arrStreams = Split(sCellValue, ",")
							For Each sStream In arrStreams
								If Not dictStreams.Exists(sStream) Then
									dictStreams.Add sStream, "P"
								End If
							Next
						End If
					End If
				Next
			End If
			
			'CHECK for ERROR value
			'CHECK output parameter references
			'CHECK unique config parameters
			
			Set oNames = oSheet.Application.ActiveWorkbook.Names
			
			'fill dictNames with cell names
			Set dictNames = CreateObject("scripting.dictionary")
			For iCounter = 1 To oNames.Count
				dictNames.Add oNames(iCounter).Name, "DEFINED"
			Next
			
			'fill dictReferences with found references
			Set dictReferences = CreateObject("scripting.dictionary")
			
			If globDebug Then
				oTraceLog.Message "chceking for reference", LOG_DEBUG
			End If
			If XL_REFERENCE > 0 Then				
				For iRow = 2 To iRowCount
					If oSheet.Cells(iRow, XL_DISABLE).Value = "" Then
						sCellValue = oSheet.Cells(iRow, XL_REFERENCE).Value
						If sCellValue <> "" Then
							If dictReferences.Exists(sCellValue) Then
								ReDim Preserve arrErrors(iErrorCount)
					    		arrErrors(iErrorCount) = "Sheet " & oSheet.Name & " [" & Chr(64 + XL_REFERENCE) & "%ROW" & iRow & "%]: Multiple occurence for reference '" & sCellValue & "'!"
					    		iErrorCount = iErrorCount + 1
							Else
								dictReferences.Add sCellValue, "DEFINED"
							End If
						End If
					End If
				Next
			End If
			
			Set dictUniqueConfigParams = CreateObject("scripting.dictionary")
			If globDebug Then
				oTraceLog.Message "checking for param and config", LOG_DEBUG
			End If
			If Not bLongCellPresent Then
				'use fast FIND method if there isn't cell longer than 1020 chars
				
				'check if any cell contains %ERROR%
				Set oCell = oSheet.UsedRange.Find("%ERROR%")
				If Not oCell Is Nothing Then
				    sFirstAddress = oCell.Address
				    Do
				    	If oSheet.Cells(oCell.Row, XL_DISABLE).Value = "" Then
					    	ReDim Preserve arrErrors(iErrorCount)
					    	arrErrors(iErrorCount) = "Sheet " & oSheet.Name & " [" & Chr(64 + oCell.Column) & "%ROW" & oCell.Row & "%]: Error value found!"
					    	iErrorCount = iErrorCount + 1
					    End If
				    	
				    	Set oCell = oSheet.UsedRange.FindNext(oCell)
				    	If Not oCell Is Nothing Then
							If oCell.Address = sFirstAddress Then Exit Do
						End If
				    Loop Until oCell Is Nothing
				End If

				'check if output params can be resolved
				Set oCell = oSheet.UsedRange.Find("param(")
				If Not oCell Is Nothing Then
					sFirstAddress = oCell.Address
				    Do
				    	'skip if disabled row
				    	If oSheet.Cells(oCell.Row, XL_DISABLE).Value = "" Then
					    	sCellValue = oCell.Value
					    	Set oRegEx = New RegExp
							oRegEx.Pattern = "param\(([a-zA-Z0-9]+)(,[a-zA-Z0-9]+)?\)"
							oRegEx.Global = True
							
							Set oMatches = oRegEx.Execute(sCellValue)
							
							For Each oMatch In oMatches
					    		If Not dictNames.Exists(oMatch.Submatches(0)) And Not dictReferences.Exists(oMatch.Submatches(0)) Then
					    			ReDim Preserve arrErrors(iErrorCount)
					    			arrErrors(iErrorCount) = "Sheet " & oSheet.Name & " [" & Chr(64 + oCell.Column) & "%ROW" & oCell.Row & "%]: Parameter reference '" & oMatch.Submatches(0) & "' not found!"
					    			iErrorCount = iErrorCount + 1
					    		End If
					    	Next
					    	
					    	Set oMatch = Nothing
					    	Set oMatches = Nothing
					    	Set oRegEx = Nothing
					    End If
				    	
				    	Set oCell = oSheet.UsedRange.FindNext(oCell)
				    	If Not oCell Is Nothing Then
							If oCell.Address = sFirstAddress Then Exit Do
						End If
				    Loop Until oCell Is Nothing
				End If

				'find unique config parameters
'				Set oCell = oSheet.UsedRange.Find("config(")
'				If Not oCell Is Nothing Then
'					sFirstAddress = oCell.Address
'				    Do
				    	'skip if disabled row
'				    	If oSheet.Cells(oCell.Row, XL_DISABLE).Value = "" Then
'					    	sCellValue = oCell.Value
'					    	Set oRegEx = New RegExp
'							oRegEx.Pattern = "config\(([^/)]+)\)"
'							oRegEx.Global = True
							
'							Set oMatches = oRegEx.Execute(sCellValue)
'							For Each oMatch In oMatches
'								If Not dictUniqueConfigParams.Exists(oMatch.Submatches(0)) Then
'									dictUniqueConfigParams.Add oMatch.Submatches(0), "X"
'								End If
'							Next
'							Set oMatches = Nothing
'						End If
				    	
'				    	Set oCell = oSheet.UsedRange.FindNext(oCell)
'				    	If Not oCell Is Nothing Then
'							If oCell.Address = sFirstAddress Then Exit Do
'						End If
'				    Loop Until oCell Is Nothing
'				End If
			Else			
				For iRow = 2 To iRowCount
					'process only enabled rows
					If oSheet.Cells(iRow, XL_DISABLE).Value = "" Then
						For iColumn = 1 To iColumnCount
							sCellValue = oSheet.Cells(iRow, iColumn).Value
							'MsgBox "Row: " & iRow & vblf & "Column: " & iColumn & vblf & "Value: " & sCellValue
							
							'if any cell contains %ERROR% then spreadsheet won't be executed
							If InStr(sCellValue, "%ERROR%") > 0 Then
								ReDim Preserve arrErrors(iErrorCount)
						    	arrErrors(iErrorCount) = "Sheet " & oSheet.Name & " [" & Chr(64 + iColumn) & "%ROW" & iRow & "%]: Error value found!"
						    	iErrorCount = iErrorCount + 1
						    End If
							
							'check if output params can be resolved
							If InStr(sCellValue, "param(") > 0 Then
								Set oRegEx = New RegExp
								oRegEx.Pattern = "param\(([a-zA-Z0-9]+)(,[a-zA-Z0-9]+)?\)"
								oRegEx.Global = True
								
								Set oMatches = oRegEx.Execute(sCellValue)
								
								For Each oMatch In oMatches
						    		If Not dictNames.Exists(oMatch.Submatches(0)) And Not dictReferences.Exists(oMatch.Submatches(0)) Then
						    			ReDim Preserve arrErrors(iErrorCount)
						    			arrErrors(iErrorCount) = "Sheet " & oSheet.Name & " [" & Chr(64 + iColumn) & "%ROW" & iRow & "%]: Parameter reference '" & oMatch.Submatches(0) & "' not found!"
						    			iErrorCount = iErrorCount + 1
						    		End If
						    	Next
						    	
						    	Set oMatch = Nothing
						    	Set oMatches = Nothing
						    	Set oRegEx = Nothing
							End If
							
							'find unique config parameters
'							If InStr(sCellValue, "config(") > 0 Then
'								Set oRegEx = New RegExp
'								oRegEx.Pattern = "config\(([^/)]+)\)"
'								oRegEx.Global = True
								
'								Set oMatches = oRegEx.Execute(sCellValue)
'								For Each oMatch In oMatches
'									If Not dictUniqueConfigParams.Exists(oMatch.Submatches(0)) Then
'										dictUniqueConfigParams.Add oMatch.Submatches(0), "X"
'									End If
'								Next
'								Set oMatches = Nothing
'							End If
							
						Next 'iColumn = 1 To iColumnCount
					End If 'oSheet.Cells(iRow, iColumn).Value = ""
				Next 'iRow = 2 To iRowCount
			End If 'Not bLongCellPresent
		
			'resolve env config file path and test set path
			'check unique config parameters if they are resolvable
'			If dictUniqueConfigParams.Count > 0 Then
				'resolve env config file path and test set path
'				If ResolveEnvConfigPathAndTestSet(sLog) = False Then
'					ReDim Preserve arrErrors(iErrorCount)
'					arrErrors(iErrorCount) = "Sheet " & oSheet.Name & ": " & sLog
'					iErrorCount = iErrorCount + 1
'				Else
					'check unique config parameters if they are resolvable
'					arrUniqueConfigParamsKeys = dictUniqueConfigParams.Keys
					
'					For Each sParameterName In arrUniqueConfigParamsKeys
'		    			If Not ResolveEnvConfigParamFromDict(sParameterName, sValueFoundSection, sValueFound) Then
'		    				ReDim Preserve arrErrors(iErrorCount)
'						    arrErrors(iErrorCount) = "Sheet " & oSheet.Name & ": config parameter not resolvable: " & sParameterName
'						    iErrorCount = iErrorCount + 1
'		    			End If
'			    	Next
'				End If
'			End If
		Next 'oSheet in oWorkbook.Worksheets

		If iErrorCount = 0 Then
			ValidateWorkbook = True
			oTraceLog.Message "Workbook is valid.",LOG_MESSAGE
		Else
			ValidateWorkbook = False
			oTraceLog.Message "Workbook is NOT valid!", LOG_ERROR
			oTraceLog.Message "Validation errors:", LOG_MESSAGE 
			'write error messages into the beginning of worksheet
			For Each sError In arrErrors
				sErrorReplaced = sError
				'add error rows count into error message
				Set oRegEx = New RegExp
				oRegEx.Pattern = "%ROW(\d+)%"
				oRegEx.Global = True
				Set oMatches = oRegEx.Execute(sError)
				If oMatches.Count > 0 Then
					For Each oMatch In oMatches
						sErrorReplaced = Replace(sErrorReplaced, oMatch, CInt(oMatch.Submatches(0)) + UBound(arrErrors) + 1)
					Next
				End If
				Set oMatch = Nothing
				Set oMatches = Nothing
				Set oRegEx = Nothing
				InsertErrorRow oWorkbook.Worksheets.Item(1), sErrorReplaced	
				oTraceLog.Message sErrorReplaced, LOG_ERROR			
			Next
		
			'activate row with errors so it's immediately visible after opening file
			oWorkbook.Worksheets(1).Activate
			oWorkbook.Worksheets(1).Rows(1).Select
			
		End If
		oTraceLog.Exited(sRoutine)
	End Function

	'==============================================================================================
	' Function/Sub:
	'==============================================================================================	
	Public Sub AssignHeaderColumnXLValues(oSheet)
		Dim iColumnCount, iCounter
		Dim sKeyword
		
		iColumnCount = oSheet.UsedRange.Columns.Count
		
		'TODO: set all XL_variables to 0 first?
		If globDebug Then
			oTraceLog.Message "Assigning columns to variables", LOG_DEBUG
		End If
		
		For iCounter = 1 To iColumnCount
			sKeyword = oSheet.Cells(1, iCounter).Text
			Select Case UCase(sKeyword)
				Case "DISABLE", "SKIP"		XL_DISABLE = iCounter
				Case "RESULT", "RES"		XL_RESULT = iCounter
				Case "STREAM", "STREAMS"	XL_STREAM = iCounter
				Case "KEYWORD", "KW"		XL_KEYWORD = iCounter
				Case "COMMENT"				XL_COMMENT = iCounter
				Case "LOG"					XL_LOG = iCounter
				Case "ID", "REFERENCE"		XL_REFERENCE = iCounter
				Case "OUTPUT_PARAMS"		XL_OUTPUT_PARAMS = iCounter
				Case "PARM_001"				XL_PARM_001 = iCounter
				Case "PARM_002"				XL_PARM_002 = iCounter
				Case "PARM_003"				XL_PARM_003 = iCounter
				Case "PARM_004"				XL_PARM_004 = iCounter
				Case "PARM_005"				XL_PARM_005 = iCounter
				Case "PARM_006"				XL_PARM_006 = iCounter
				Case "PARM_007"				XL_PARM_007 = iCounter
				Case "PARM_008"				XL_PARM_008 = iCounter
				Case "PARM_009"				XL_PARM_009 = iCounter
				Case "PARM_010"				XL_PARM_010 = iCounter
				
				'compatibility with older spreadsheets
				Case "OUTPUT"				XL_LOG = iCounter
				Case "MESSAGE"				XL_OUTPUT_PARAMS = iCounter
			End Select
		Next
		
		'For recovery
		oQRSDataObject.QRSXL_RESULT = XL_RESULT
		oQRSDataObject.QRSXL_KEYWORD = XL_KEYWORD
		oQRSDataObject.QRSXL_COMMENT = XL_COMMENT
		oQRSDataObject.QRSXL_LOG = XL_LOG
		oQRSDataObject.QRSXL_REFERENCE = XL_REFERENCE
		oQRSDataObject.QRSXL_OUTPUT_PARAMS = XL_OUTPUT_PARAMS
		oQRSDataObject.QRSXL_PARM_001 = XL_PARM_001
		oQRSDataObject.QRSXL_PARM_002 = XL_PARM_002
		oQRSDataObject.QRSXL_PARM_003 = XL_PARM_003
		oQRSDataObject.QRSXL_PARM_004 = XL_PARM_004
		oQRSDataObject.QRSXL_PARM_005 = XL_PARM_005
		oQRSDataObject.QRSXL_PARM_006 = XL_PARM_006
		oQRSDataObject.QRSXL_PARM_007 = XL_PARM_007
		oQRSDataObject.QRSXL_PARM_008 = XL_PARM_008
		oQRSDataObject.QRSXL_PARM_009 = XL_PARM_009
		oQRSDataObject.QRSXL_PARM_010 = XL_PARM_010
	End Sub
	
	'==============================================================================================
	' Function/Sub:
	'==============================================================================================	
	Public Sub InsertErrorRow(oSheet, sError)
		Dim sCellValue
		
		oSheet.Rows(2).Insert
		oSheet.Rows(2).Font.ColorIndex		= 2
    	oSheet.Rows(2).Font.Size			= 11
		oSheet.Rows(2).Interior.ColorIndex	= 3
		
'		oSheet.Rows(2).Select
'    	oSheet.Application.Selection.Font.ColorIndex		= 2
'    	oSheet.Application.Selection.Font.Size				= 11
'		oSheet.Application.Selection.Interior.ColorIndex	= 3
		
		'oSheet.Rows(2).Cells(1, 1).WrapText = False
		oSheet.Rows(2).Cells(1, 1).Value = "Validation error: " & sError
		oSheet.Rows(2).Cells(1, 1).WrapText = False
	End Sub

	'==============================================================================================
	' Function/Sub:
	'==============================================================================================
	Public Function DataAndConfigParamsCellProcessing(byVal sString, sLog)
		Dim oRegExConfig, oRegExData, oCell, oRegEx, oMatches, oMatch
		Dim sRegexpConfig, sRegexpData, sReplacePattern, sCellValue, sResourceName, sValueFound, sLookupColumn, sLookupValue, sResultColumn
		Dim bLongCellPresent
		Dim arrParams, arrIndexNameValuePair
		
		If globDebug Then
			Const sRoutine = "clsVBSFramework.DataAndConfigParamsCellProcessing"
			oTraceLog.Entered(sRoutine)
		End If
		
		
		'return from recursive call if error occured
		If sString = "%ERROR%" Then Exit Function
	
		'define regular expressions for matching data() and config() parameters
		sRegexpConfig = "config\(([^\(\)]+)\)"
		sRegexpData = "data\(([^\(\)]+)\)"
		Set oRegExConfig = New RegExp
		Set oRegExData = New RegExp
		oRegExConfig.Pattern = sRegexpConfig
		oRegExConfig.Global = True
		oRegExData.Pattern = sRegexpData
		oRegExData.Global = True
	
		'CONFIG parameters - config(sLookupValue)
		For Each oMatch In oRegExConfig.Execute(sString)
			
			sValueFound = ""
			sReplacePattern = "config(" & oMatch.Submatches(0) & ")"
			
			'hardcoded resource name for config() parameters
			sResourceName = "ENV_CONFIGURATION"
			
			'pre-cache content of external data spreadsheet into memory if not already done
			If Not dictExternalData.Exists(sResourceName) Then
			
			'	PrecacheExcelToDictionary sResourceName, sLog
				PrecacheExcelToDictionary sResourceName
				'error if pre-caching failed
				If Not dictExternalData.Exists(sResourceName) Then
					
					sLog = sLog & "E: Resource file '" & sResourceName & "' not found neither locally nor in HPQC Resources. Check file name, file path or HPQC connection." & vblf
					oTraceLog.Exited(sRoutine)
					sString = "%ERROR%" : Exit Function
				End If
			End If
			
			If bDataSetPrompt Then
				'Set oQC = New clsQC
				'resolve execution path
				If oQC.IsQCRun() Then
					sDataSetPath = globCurrentTSTest.TestSet.TestSetFolder.Path & "\" & globCurrentTSTest.TestSet.Name
				Else
					'display popup if test run outside of QC
						'TODO create proper handling for local QTP test, not only CMD test
'					If Not globRunMode = CMD_TEST Then
'						sDataSetPath = InputBox("The HPQC/ALM Test Lab path/folder indicates the DataSet to use for test data/config values." & vblf & vblf & "This test is not running from HPQC/ALM. Please enter a path/folder to indicate which DataSet to use." & vblf & vblf & "Examples:" & vblf & vblf & "ENV1" & vblf & "Root\Folder\SubFolder", "Enter DataSet", sDataSetPath)
'					Else 						
						'set folder for DataSet according to globDataPath (e.g default is spreadsheets)
'						sDataSetPath = oFso.GetBaseName(globDataPath)
'					End If
					If 	Not oFso.FolderExists(globDataPath) Then
							oTraceLog.Message "Path to Data folder does not exists", LOG_ERROR
							oTraceLog.Exited(sRoutine)
							DataAndConfigParamsCellProcessing = "%ERROR%" : Exit Function
					End If
							
					sDataSetPath = Mid(globDataPath,InStr(globDataPath,":") + 2)
				End If
				bDataSetPrompt = False
			'	Set oQC = Nothing
			End If
			
'			If sQCExecutionPath = "" Then
'				Set oQC = New clsQC
				'resolve execution path
'				If oQC.IsQCRun() Then
'					sQCExecutionPath = globCurrentTSTest.TestSet.TestSetFolder.Path & "\" & globCurrentTSTest.TestSet.Name
'				Else
					'display popup if test run outside of QC
'					sQCExecutionPath = InputBox("The HPQC/ALM Test Lab path/folder indicates the DataSet to use for test data/config values." & vblf & vblf & "This test is not running from HPQC/ALM. Please enter a path/folder to indicate which DataSet to use." & vblf & vblf & "Examples:" & vblf & vblf & "ENV1" & vblf & "Root\Folder\SubFolder", "Enter DataSet", sDataSetPath)
'				End If
'				Set oQC = Nothing
'			End If
			
			'resolve parameter value
			sValueFound = ResolveExternalData(sResourceName, "1", oMatch.Submatches(0), "2", sDataSetPath, sLog)			    			
			
			If sValueFound = "%NOT FOUND%" Then
				sLog = sLog & "E: Can't resolve config("  & oMatch.Submatches(0) & ") value." & vbLf
				oTraceLog.Exited(sRoutine)
				DataAndConfigParamsCellProcessing = "%ERROR%" : Exit Function
    		Else
    			sLog = sLog & "D: Parameter '" & oMatch.Submatches(0) & "' replaced with value '" & sValueFound & "'" & vblf
				sString = Replace(sString, sReplacePattern, sValueFound)
    		End If
		Next
	
		'DATA parameters - data(sResourceName, [sLookupColumn=]sLookupValue[, sResultColumn])
		For Each oMatch In oRegExData.Execute(sString)
			sValueFound = ""
			sReplacePattern = "data(" & oMatch.Submatches(0) & ")"
			arrParams = Split(oMatch.Submatches(0), ",")

    		If UBound(arrParams) = 1 Or UBound(arrParams) = 2 Then
    			arrIndexNameValuePair = Split(arrParams(1), "=")
    			
    			If UBound(arrIndexNameValuePair) = 1 Then
    				sLookupColumn = arrIndexNameValuePair(0)
    				sLookupValue = arrIndexNameValuePair(1)
    			Else
    				sLookupColumn = "1"
    				sLookupValue = arrParams(1)
    			End If
    			
    			sResourceName = UCase(Trim(arrParams(0)))
    			
				'pre-cache content of external data spreadsheet into memory if not already done
    			If Not dictExternalData.Exists(sResourceName) Then
    			'	PrecacheExcelToDictionary sResourceName, sLog
					PrecacheExcelToDictionary sResourceName

    				'error if pre-caching failed
    				If Not dictExternalData.Exists(sResourceName) Then
    					sLog = sLog & "E: Resource file '" & sResourceName & "' not found neither locally nor in HPQC Resources. Check file name, file path or HPQC connection." & vblf
					
    					oTraceLog.Exited(sRoutine)
    					DataAndConfigParamsCellProcessing = "%ERROR%" : Exit Function
    				End If
    			End If
    			
    			If UBound(arrParams) = 1 Then
    				sResultColumn = "2"
    			Else
    				sResultColumn = arrParams(2)
    			End If
    			
    			If bDataSetPrompt Then
					'Set oQC = New clsQC
					'resolve execution path
					If oQC.IsQCRun() Then
						sDataSetPath = globCurrentTSTest.TestSet.TestSetFolder.Path & "\" & globCurrentTSTest.TestSet.Name
					Else
				
						'display popup if test run outside of QC
					
						'TODO create proper handling for local QTP test, not only CMD test
						'If Not globRunMode = CMD_TEST Then
						'	sDataSetPath = InputBox("The HPQC/ALM Test Lab path/folder indicates the DataSet to use for test data/config values." & vblf & vblf & "This test is not running from HPQC/ALM. Please enter a path/folder to indicate which DataSet to use." & vblf & vblf & "Examples:" & vblf & vblf & "ENV1" & vblf & "Root\Folder\SubFolder", "Enter DataSet", sDataSetPath)
						'Else 						
							'set folder for DataSet according to globDataPath (e.g default is spreadsheets)
						If 	Not oFso.FolderExists(globDataPath) Then
							oTraceLog.Message "Path to Data folder does not exists", LOG_ERROR
							oTraceLog.Exited(sRoutine)
							DataAndConfigParamsCellProcessing = "%ERROR%" : Exit Function
						End If
							
						sDataSetPath = Mid(globDataPath,InStr(globDataPath,":") + 2)
					
					'	End If
						
					End If
					bDataSetPrompt = False
					'Set oQC = Nothing
				End If
    			
'    			If sQCExecutionPath = "" Then
    				'resolve execution path
'    				Set oQC = New clsQC
'    				If oQC.IsQCRun() Then
'						sQCExecutionPath = globCurrentTSTest.TestSet.TestSetFolder.Path & "\" & globCurrentTSTest.TestSet.Name
'					Else
						'display popup if test run outside of QC
'						sQCExecutionPath = InputBox("The HPQC/ALM Test Lab path/folder indicates the DataSet to use for test data/config values." & vblf & vblf & "This test is not running from HPQC/ALM. Please enter a path/folder to indicate which DataSet to use." & vblf & vblf & "Examples:" & vblf & vblf & "ENV1" & vblf & "Root\Folder\SubFolder", "Enter DataSet", sDataSetPath)
'					End If
'					Set oQC = Nothing
'				End If
    			
    			'resolve parameter value
    			sValueFound = ResolveExternalData(sResourceName, sLookupColumn, sLookupValue, sResultColumn, sDataSetPath, sLog)			    			
    			
    			If sValueFound = "%NOT FOUND%" Then
    				sLog = sLog & "E: Can't resolve data("  & oMatch.Submatches(0) & ") value." & vbLf
    				oTraceLog.Exited(sRoutine)
    				DataAndConfigParamsCellProcessing = "%ERROR%" : Exit Function
				End If
    		Else
    			sLog = sLog & "E: Unsupported usage of data() keyword. Usage: data(ResourceName, [LookupColumnName=]LookupValue[, ResultColumnName])." & vblf
    			oTraceLog.Exited(sRoutine)
    			DataAndConfigParamsCellProcessing = "%ERROR%" : Exit Function
    		End If
			
			sLog = sLog & "D: Data(" & oMatch.Submatches(0) & ") replaced with value: " & sValueFound & vblf
			sString = Replace(sString, sReplacePattern, sValueFound)
		Next

		'recursive function call
		If (oRegExConfig.Execute(sString).Count > 0) Or (oRegExData.Execute(sString).Count > 0) Then
			sString = DataAndConfigParamsCellProcessing(sString, sLog)
		End If

		DataAndConfigParamsCellProcessing = sString
		If globDebug Then
			oTraceLog.Exited(sRoutine)
		End If	
	End Function

	'==============================================================================================
	' Function/Sub:
	'==============================================================================================
	Public Function DataAndConfigParamsRowProcessing(oRow, sLog)
		Dim iRetVal, iColumn
		Dim oCell, oRegEx, oMatches, oMatch
		Dim sCellValue, sResourceName, sValueFound, sReplacePattern, sFirstAddress', sLog
		Dim bLongCellPresent
		
		If globDebug Then
			Const sRoutine = "clsVBSFramework.DataAndConfigParamsRowProcessing"
			oTraceLog.Entered(sRoutine)
		End If
		
		iRetVal = True
		bLongCellPresent = False

		For Each oCell In oRow.Range(oRow.Cells(1, XL_PARM_001), oRow.Cells(1, XL_PARM_001 + 10))
			If Len(oCell) > 1020 Then bLongCellPresent = True
		Next
		
		'quick check if cells contain config() or data()
		Set oCell = oRow.Worksheet.Range(oRow.Cells(1, XL_PARM_001), oRow.Cells(1, XL_PARM_001 + 10)).Find("config(")
		If oCell Is Nothing Then
			Set oCell = oRow.Worksheet.Range(oRow.Cells(1, XL_PARM_001), oRow.Cells(1, XL_PARM_001 + 10)).Find("data(")
		End If
		If oCell Is Nothing And Not bLongCellPresent Then
			'skip row
			If globDebug Then
			'WriteIntoLogAndOutputParamsCells oRow, "D: skipping resolving of data() and config() params." & vblf, ""
				oTraceLog.Message "D: skipping resolving of data() and config() params.", LOG_DEBUG
			End If
			Set oCell = Nothing
			if globDebug Then
				oTraceLog.Exited(sRoutine)
			End If
			DataAndConfigParamsRowProcessing = True : Exit Function
		End If
		Set oCell = Nothing
		
		'loop through all input parameter cells
		For iColumn = XL_PARM_001 To XL_PARM_001 + 10
			sCellValue = oRow.Cells(1, iColumn).Value
			'MsgBox "Row: " & iRow & vblf & "Column: " & iColumn & vblf & "Value: " & sCellValue
		'	If globDebug Then
		'		oTraceLog.Message "Row: " & iRow & " Column: " & iColumn  & " Value: " & sCellValue, LOG_DEBUG
		'	End If
			
			sValueFound = DataAndConfigParamsCellProcessing(sCellValue, sLog)
			
			If sValueFound = "%ERROR%" Then
				
			'	WriteIntoLogAndOutputParamsCells oRow, sLog, ""
			'	oTraceLog.Message sLog, LOG_ERROR
				if globDebug Then
					oTraceLog.Exited(sRoutine)
				End If
					
				DataAndConfigParamsRowProcessing = False : Exit Function
			ElseIf sValueFound = sCellValue Then
				'do nothing
				If globDebug Then
					oTraceLog.Message "Found value is same as Cell value", LOG_DEBUG
				End If
			Else
				'replace data() with value
				If globDebug Then
					oTraceLog.Message "Replacing values...", LOG_DEBUG
				End If
				oRow.Cells(1, iColumn).NumberFormat = "@"
				oRow.Cells(1, iColumn).Value = sValueFound
				oRow.Cells(1, iColumn).NumberFormat = "General"
			End If
		Next 'iColumn = XL_PARM_001 To XL_PARM_001 + 10
		
		'write logging messages into log cell - change		
	'	WriteIntoLogAndOutputParamsCells oRow, sLog, ""
	'	oTraceLog.Message sLog, LOG_MESSAGE
		DataAndConfigParamsRowProcessing = iRetVal
	
		If globDebug Then
			oTraceLog.Exited(sRoutine)
		End If	
	End Function
	
	'==============================================================================================
	' Function/Sub:
	'==============================================================================================
	Public Sub ProcessOutputParams(oRow, sLog)
		Dim dictOutputParams
		Dim arrKeys, arrItems 
		Dim oApp, oReplaceRange, oSheet, oCell, oNamedCell
		Dim iRow, iColumn, iLen, iRowCount, iColumnCount, iCounter
		Dim sReplacePattern, sIDC, sCellValue, sReference, sFirstAddress
		Dim bNamedCell, bLongCellPresent
		If globDebug Then
			oTraceLog.Message "Processing outputparams...", LOG_DEBUG
		End if
		bNamedCell = False
		
		Set oApp = oRow.Application
		Set oSheet = oApp.ActiveSheet
		iRowCount = oSheet.UsedRange.Rows.Count
		iColumnCount = oSheet.UsedRange.Columns.Count

		sCellValue = oRow.Cells(1, XL_OUTPUT_PARAMS).Value
		
		'get rid of rightmost comma character
		If Right(sCellValue, 1) = "," Then
			sCellValue = Left(sCellValue, Len(sCellValue) - 1)
			oRow.Cells(1, XL_OUTPUT_PARAMS).Value = sCellValue
		End If
		
		'read output parameter values
		Set dictOutputParams = CreateObject("Scripting.Dictionary")
		
		'support for multiple parameters divided with comma
		SimpleDictKeyValuePairs dictOutputParams, sCellValue, ",", "="
		
		arrKeys = dictOutputParams.Keys : arrItems = dictOutputParams.Items
	
		'change ~ character to = (to support = character in SetParameter)
		'change | character to , (to support , character in SetParameter)
		For iCounter = 0 To UBound(arrItems)
			arrItems(iCounter) = Replace(Replace(arrItems(iCounter), "^", ","), "~", "=")
		Next
		
		'TODO: case-insensitive replace
		
		If dictOutputParams.Count > 0 Then
			sReference = ""
			'TODO: make sure XL_REFERENCE is reset to 0 before each sheet processing
			If XL_REFERENCE > 0 Then
				sReference = oRow.Cells(1, XL_REFERENCE).Value
			End If
			'if reference not available, try to use cell name
			If sReference = "" Then
				On Error Resume Next
				'.Name raises an error if it's not defined
				Set oNamedCell = oRow.Cells(1, XL_OUTPUT_PARAMS).Name
				If Err.Number = 0 Then
					sReference = oNamedCell.Name
				End If
				On Error Goto 0
				Set oNamedCell = Nothing
			End If
			'MsgBox sReference
			
			'check if reference has been found
			If sReference = "" Then
				sLog = sLog & "W: reference column missing/empty and output params cell is not named => no processing of output parameters."  & vblf
			Else
				oApp.Calculation = -4135 'manual calculation
				
				bLongCellPresent = False
				For Each oCell In oSheet.UsedRange
					If Len(oCell) > 1020 Then bLongCellPresent = True
				Next
				
				If Not bLongCellPresent Then
					'use fast FIND method if there isn't cell longer than 1020 chars
					
					'replace shortened version: param(CELLNAME)
					sReplacePattern = "param(" & sReference & ")"

					Set oCell = oSheet.UsedRange.Find(sReplacePattern)
					If Not oCell Is Nothing Then
						sFirstAddress = oCell.Address
					    Do
					        sCellValue = oSheet.Cells(oCell.Row, oCell.Column).Value
					        sCellValue = Replace(sCellValue, sReplacePattern, arrItems(0), 1, -1, 1)
					        oSheet.Cells(oCell.Row, oCell.Column).NumberFormat = "@"
					        oSheet.Cells(oCell.Row, oCell.Column).Value = sCellValue
					        oSheet.Cells(oCell.Row, oCell.Column).NumberFormat = "General"
					        sLog = sLog & "I: output param replaced [" & Chr(64 + oCell.Column) & oCell.Row & "]" & vblf
					        Set oCell = oSheet.UsedRange.FindNext(oCell)
					        If Not oCell Is Nothing Then
								If oCell.Address = sFirstAddress Then Exit Do
							End If
					    Loop Until oCell Is Nothing
					End If
					
					'replace full version "param(CELLNAME,paramname)"
					For iCounter = 0 To dictOutputParams.Count - 1
						sReplacePattern = "param(" & sReference & "," & arrKeys(iCounter) & ")"
						
						Set oCell = oSheet.UsedRange.Find(sReplacePattern)
						If Not oCell Is Nothing Then
							sFirstAddress = oCell.Address
						    Do
						        sCellValue = oSheet.Cells(oCell.Row, oCell.Column).Value
						        sCellValue = Replace(sCellValue, sReplacePattern, arrItems(iCounter), 1, -1, 1)
						        oSheet.Cells(oCell.Row, oCell.Column).NumberFormat = "@"
						        oSheet.Cells(oCell.Row, oCell.Column).Value = sCellValue
						        oSheet.Cells(oCell.Row, oCell.Column).NumberFormat = "General"
						        sLog = sLog & "I: output param replaced [" & Chr(64 + oCell.Column) & oCell.Row & "]" & vblf
						        Set oCell = oSheet.UsedRange.FindNext(oCell)
						        If Not oCell Is Nothing Then
									If oCell.Address = sFirstAddress Then Exit Do
								End If
						    Loop Until oCell Is Nothing
						End If
					Next
				Else
					'use slow ONE-BY-ONE method if there is cell longer than 1020 chars
					
					'replace shortened version: param(CELLNAME)
					sReplacePattern = "param(" & sReference & ")"
					
					For iRow = oRow.Row + 1 To iRowCount
						If oSheet.Cells(iRow, XL_DISABLE).Value = "" Then
							For iColumn = XL_PARM_001 To XL_PARM_001 + 10
								'iLen = Len(oSheet.Cells(iRow, iColumn))
								'If Len(oSheet.Cells(iRow, iColumn)) = 0 And Len(oSheet.Cells(iRow, iColumn + 1)) = 0 Then
								If Not IsNull(oSheet.range( oSheet.Cells(iRow, iColumn),  oSheet.Cells(iRow, XL_PARM_001 + 10) ).FormulaArray ) Then
									'move to next row if two subsequent cells are empty
									Exit For
								Else
									sCellValue = oSheet.Cells(iRow, iColumn).Value
									If InStr(1, sCellValue, sReplacePattern, 1) > 0 Then
										sCellValue = Replace(sCellValue, sReplacePattern, arrItems(0), 1, -1, 1)
										oSheet.Cells(iRow, iColumn).NumberFormat = "@"
										oSheet.Cells(iRow, iColumn).Value = sCellValue
										oSheet.Cells(iRow, iColumn).NumberFormat = "General"
										sLog = sLog & "I: output param replaced [" & Chr(64 + iColumn) & iRow & "]" & vblf
									End If
								End If
							Next
						End If
					Next

					'replace full version "param(CELLNAME,paramname)"
					For iCounter = 0 To dictOutputParams.Count - 1
						sReplacePattern = "param(" & sReference & "," & arrKeys(iCounter) & ")"
						
						For iRow = oRow.Row + 1 To iRowCount
							If oSheet.Cells(iRow, XL_DISABLE).Value = "" Then
								For iColumn = XL_PARM_001 To XL_PARM_001 + 10
									'iLen = Len(oSheet.Cells(iRow, iColumn))
									'If Len(oSheet.Cells(iRow, iColumn)) = 0 And Len(oSheet.Cells(iRow, iColumn + 1)) = 0 Then
									If Not IsNull(oSheet.range( oSheet.Cells(iRow, iColumn),  oSheet.Cells(iRow, XL_PARM_001 + 10) ).FormulaArray ) Then
										'move to next row if two subsequent cells are empty
										Exit For
									Else
										sCellValue = oSheet.Cells(iRow, iColumn).Value
										If InStr(1, sCellValue, sReplacePattern, 1) > 0 Then
											sCellValue = Replace(sCellValue, sReplacePattern, arrItems(iCounter), 1, -1, 1)
											oSheet.Cells(iRow, iColumn).NumberFormat = "@"
											oSheet.Cells(iRow, iColumn).Value = sCellValue
											oSheet.Cells(iRow, iColumn).NumberFormat = "General"
											sLog = sLog & "I: output param replaced [" & Chr(64 + iColumn) & iRow & "]" & vblf
										End If
									End If
								Next
							End If
						Next
					Next
				End If 'bLongCellPresent
				oApp.Calculation = -4105 'automatic calculation
				
			End If 'sReference = ""
		End If 'dictOutputParams.Count > 0
		
		Set dictOutputParams = Nothing
		Set oSheet = Nothing
		
	End Sub
	
'==============================================================================================
' End Class clsVBSFramework
'==============================================================================================
End Class

'==================================================================================================
' Class Start
'==================================================================================================

Class clsTestObjects

	'==============================================================================================
	' CLASS VARIABLES
	'==============================================================================================

	'----------------------------------------------------------------------------------------------
	' Public Variables
	'----------------------------------------------------------------------------------------------

	'----------------------------------------------------------------------------------------------
	' Private Variables
	'----------------------------------------------------------------------------------------------

	Private dictTestObject		
	Private dictFile			
	'==============================================================================================
	' CLASS PROPERTIES
	'==============================================================================================

	'==============================================================================================
	' CLASS SUBS & FUNCTIONS
	'==============================================================================================

	'----------------------------------------------------------------------------------------------
	' CLASS_INITIALIZE
	'----------------------------------------------------------------------------------------------
	Private Sub Class_Initialize
		Set dictTestObject = createobject("scripting.dictionary")
		Set dictFile = createobject("scripting.dictionary")
	End Sub

	'----------------------------------------------------------------------------------------------
	' CLASS_TERMINATE
	'----------------------------------------------------------------------------------------------
	Private Sub Class_Terminate
		Set dictTestObject = Nothing
		Set dictFile = Nothing
	End Sub

	'==============================================================================
	' Function:		IsFileLoaded
	' Purpose:
	' Parameters:	sFile - name
	' Returns:		True/False
	'==============================================================================
	Public function IsFileLoaded(sFile)
		if dictFile.exists(sFile) then		'Is sTestObject there?
			IsFileLoaded = True				'Is there
		else
			IsFileLoaded = False			'Not there.
		end if
	End function

	'==============================================================================
	' Function:		AddFile
	' Purpose:		Set the application object in dictTestObject
	' Parameters:	sFile - name
	' Returns:		True/False
	'==============================================================================
	Public function AddFile(sFile)
		if IsFileLoaded(sFile) = True then			'Is sTestObject there?
			AddFile = False							'Already there
		else
			dictFile.add sFile, True				'Not there. Add it to dictTestObject
			AddFile = True
		end if
	End function

	'==============================================================================
	' Function:		IsLoaded
	' Purpose:
	' Parameters:	sTestObject - name
	' Returns:		True/False
	'==============================================================================
	Public function IsLoaded(sTestObject)
		if dictTestObject.exists(sTestObject) then		'Is sTestObject there?
			IsLoaded = True						'Is there
		else
			IsLoaded = False					'Not there.
		end if
	End function

	'==============================================================================
	' Function:		Add
	' Purpose:		Set the application object in dictTestObject
	' Parameters:	sTestObject - name
	' Returns:		Application Object
	'==============================================================================
	Public function Add(sTestObject, oTestObject)
		if IsLoaded(sTestObject) = True then			'Is sTestObject there?
			Set Add = Nothing					'Already there
		else
			dictTestObject.add UCase(sTestObject), oTestObject		'Not there. Add it to dictTestObject
			Set Add = oTestObject
		end if
	End function

	'==============================================================================
	' Function:		Item
	' Purpose:		Get the application object from dictTestObject
	' Parameters:	sTestObject - name
	' Returns:		Application Object
	'==============================================================================
	Public function Item(sTestObject)
		if IsLoaded(sTestObject) = True then					'Is sTestObject there?
			Set Item = dictTestObject.item(sTestObject)			'Return it
		else
			Set Item = Nothing							'Not there. Return Nothing.
		end if
	End function

'==============================================================================================
' End Class clsTestObjects
'==============================================================================================
End Class


'==================================================================================================
' Class Start
'==================================================================================================

Class clsTraceLog
	'==============================================================================================
	' CLASS VARIABLES
	'==============================================================================================

	'----------------------------------------------------------------------------------------------
	' Public Variables
	'----------------------------------------------------------------------------------------------

		
	'----------------------------------------------------------------------------------------------
	' Private Variables
	'----------------------------------------------------------------------------------------------
'	Private bOn
'	Public ParentName
	Public LogFilePath

	'==============================================================================================
	' CLASS PROPERTIES
	'==============================================================================================

	'==============================================================================================
	' CLASS SUBS & FUNCTIONS
	'==============================================================================================

	'----------------------------------------------------------------------------------------------
	' CLASS_INITIALIZE
	'----------------------------------------------------------------------------------------------
	Private Sub Class_Initialize
	'	ParentName = ""
	'	bOn = False
		gTraceLogInstanceNumber = gTraceLogInstanceNumber + 1
		LogFilePath = PATH_HOSTLOGFILE 'use it by default
	End Sub


	'----------------------------------------------------------------------------------------------
	' CLASS_TERMINATE
	'----------------------------------------------------------------------------------------------
	'==============================================================================
	' Function:		TurnOn
	' Purpose:
	' Parameters:
	'==============================================================================
	Public Sub TurnOn()
'		bOn = True
		gTraceLogOn = True		
		AppendToFile LogFilePath, Now() & " TraceLog instance " & gTraceLogInstanceNumber & " turned on for All." & vblf
			
	End Sub
	'==============================================================================
	' Function:		TurnOff
	' Purpose:
	' Parameters:
	'==============================================================================
	Public Sub TurnOff()
'		bOn = False
		gTraceLogOn = False
		gTraceLogDepth = 0		
		AppendToFile LogFilePath, Now() & " TraceLog instance " & gTraceLogInstanceNumber & " turned off for All." & vblf 		
	End Sub

	'==============================================================================
	' Function:		TurnOn
	' Purpose:
	' Parameters:
	'==============================================================================
'	Public Sub TurnOnForMe()
'		bOn = True
'		Select Case globRunMode
'			Case QTP_TEST,QTP_LOCAL_TEST
'				Reporter.ReportEvent micDone, "TraceLog instance " & gTraceLogInstanceNumber & " turned on for " & Parentname & ".", _
'										"TraceLog instance " & gTraceLogInstanceNumber & " turned on for " & Parentname & "."			
'			Case VAPI_XP_TEST
'				TDHelper.AddStepToRun "TraceLog instance " & gTraceLogInstanceNumber & " turned on for " & ParentName & ".", _
'								"TraceLog instance " & gTraceLogInstanceNumber & " turned on for " & ParentName & "."
'			Case CMD_TEST
'				AppendToFile PATH_RESOURCES & "Trace.Log", sTimeStamp & " TraceLog instance " & gTraceLogInstanceNumber & " turned on for " & ParentName & "." & vblf
'		End Select
		'AppendToFile PATH_RESOURCES & "Trace.Log", sTimeStamp & " TraceLog instance " & gTraceLogInstanceNumber & " turned on for " & ParentName & "." & vblf
'		AppendToFile LogFilePath, sTimeStamp & " TraceLog instance " & gTraceLogInstanceNumber & " turned on for " & ParentName & "." & vblf
				
'	End Sub

	'==============================================================================
	' Function:		TurnOff
	' Purpose:
	' Parameters:
	'==============================================================================
'	Public Sub TurnOffForMe()
'		bOn = False
'		Select Case globRunMode
'			Case QTP_TEST,QTP_LOCAL_TEST	
'				Reporter.ReportEvent micDone, "TraceLog instance " & gTraceLogInstanceNumber & " turned off for " & Parentname & ".", _
'										"TraceLog instance " & gTraceLogInstanceNumber & " turned off for " & Parentname & "."		
'			Case VAPI_XP_TEST
'				TDHelper.AddStepToRun "TraceLog instance " & gTraceLogInstanceNumber & " turned off for " & ParentName & ".", _
'											"TraceLog instance " & gTraceLogInstanceNumber & " turned off for " & ParentName & "."
'			Case CMD_TEST
'					AppendToFile PATH_RESOURCES & "Trace.Log", Now() & " TraceLog instance " & gTraceLogInstanceNumber & " turned off for " & ParentName & "." & vblf 
'		End Select	
	'	AppendToFile PATH_RESOURCES & "Trace.Log", Now() & " TraceLog instance " & gTraceLogInstanceNumber & " turned off for " & ParentName & "." & vblf 
'		AppendToFile LogFilePath, Now() & " TraceLog instance " & gTraceLogInstanceNumber & " turned off for " & ParentName & "." & vblf 
			
'	End Sub

	'==============================================================================
	' Function:		Entered
	' Purpose:
	' Parameters:	Routine name
	'==============================================================================
	Public Sub Entered(sRoutine)
		Dim sInArrows		
	
	'	If gTraceLogOn = True Or bOn = True Then
		If gTraceLogOn = True then
			gTraceLogDepth = gTraceLogDepth + 1
			sInArrows = mid("> > > > > > > > > > ", 1, gTraceLogDepth*2)
			
			AppendToFile LogFilePath, Now() & " " & sInArrows & "> ENTERED: " & sRoutine & vblf
		End If		
	End Sub

	'==============================================================================
	' Function:		Message
	' Purpose:
	' Parameters:	Array - contains in order> Message, Status of Message
	'               possible TYPES:
	'				LOG_ERROR = "ERROR"
	'				LOG_WARNING = "WARNING"
	'				LOG_SLOG = "LOG"
	'				LOG_MESSAGE = ""	
	'				String - for backcompatibility message
	'
	'==============================================================================
	Public Sub Message(sText, sType)
		
		If gTraceLogOn = True then	
			If sText <> "" Then
				If (sType <> "")	Then					
					AppendToFile LogFilePath, Now() & " " &   sType & " " & sText & vblf
				Else
					AppendToFile LogFilePath, Now()  & " " & sText  & vblf
				End If	
			End If
		End If	
	End Sub
	'==============================================================================
	' Function:		HostLogMessage
	' Purpose:		Message for Automatin.log, for Host Trace log
	' Parameters:	Message
	' Author		
	'==============================================================================
	Public Sub HostLogMessage(sMessage)		
		If gTraceLogOn = True then		
			AppendToFile PATH_HOSTLOGFILE,  YYYYMMDDHHMMSS(dFrameworkStart) & " " & Now() & " " & sMessage & vblf
		End If		
	End Sub
	'==============================================================================
	' Function:		Exited
	' Purpose:
	' Parameters:	Routine name
	'==============================================================================
	Public Sub Exited(sRoutine)
		Dim sOutArrows

	'	If gTraceLogOn = True Or bOn = True Then
		If gTraceLogOn = True then
			gTraceLogDepth = gTraceLogDepth - 1
			sOutArrows = mid("< < < < < < < < < < ", 1, (gTraceLogDepth+1) * 2)		

			AppendToFile LogFilePath, Now() & " " & sOutArrows & "< EXITED: " & sRoutine & vblf								
		End If
	End Sub
	
	'==============================================================================
	' Sub:		StepMessage
	' Purpose:		Report of test in steps in ALM
	' Parameters:	sStepName = name of step> TestObject.Action (keyword)
	'				iRetVal = return value of step in RowDispach
	'				sLog = log rom action
	'				sInputParams = input params from action
	'				sOutputParams = output params from action
	' Author		
	'==============================================================================
	Public Sub StepMessage(sStepName, iRetVal, sLog, arrInputParams, sOutputParams)		
		Dim sStatus 'for VAPI report
		Dim sDescription ' description contains input params, output params and log
		Dim kStatus 'for QTP report	
		Dim arrOutputParams
		Dim i	
				
		If sStepName <> "" Then	'skiping blank lines	
			Select Case iRetVal
				Case XL_DISPATCH_END
					sStatus = "N/A"
					kStatus = 2 'micDone
				Case XL_DISPATCH_SKIP
					sStatus = "Not Run" 'added for step to not be run
					kStatus = 2 'micDone
				Case XL_DISPATCH_PASS
					sStatus = "Passed"
					kStatus = 0 'micPass
				Case XL_DISPATCH_FAIL
					sStatus = "Failed"
					kStatus = 1 'micFail
				Case XL_DISPATCH_CANCEL
					sStatus = "N/A"
					kStatus = 3 'micWarning
				Case XL_DISPATCH_FAILCONTINUE
					sStatus = "Failed"
					kStatus = 1 'micFail
				Case XL_DISPATCH_UNKNOWN 'fail step if its unknown keyword
					sStatus = "Failed"
					kStatus = 1 'micFail
				Case Else	
					sStatus = "N/A"
					kStatus = 3 'micWarning							
			End Select	
					
			sDescription = "LOG:" & vblf & sLog  & vblf 
			
			sDescription =  sDescription & "INPUT PARAMS:" & vblf 
			
			If IsArray(arrInputParams) Then
				For i = 0 To UBound(arrInputParams) - 1
					If arrInputParams(i) <> "" Then
						sDescription = sDescription & "PARAM_" & i + 1 & "=" & arrInputParams(i)  & vblf
					End If
				Next
			End If
			
			sDescription =  sDescription & vblf & "OUTPUT PARAMS:" & vblf 
		
			If sOutputParams <> "" Then
				arrOutputParams = Split(sOutputParams,",",-1,1)
				
				For i = 0 To UBound(arrOutputParams)	
					If arrOutputParams(i) <> "" Then			 		
						sDescription = sDescription & 	Replace(Replace(arrOutputParams(i),  "^", ","), "~", "=")  & vblf
					End If
				Next
			End If
					
			
		'	If gTraceLogOn = True Or bOn = True Then
			
			Select Case globRunMode
					Case QTP_TEST,QTP_LOCAL_TEST						
						Reporter.ReportEvent kStatus, sStepName, sDescription			
					Case VAPI_XP_TEST
						'Params for AddStepToRun(Name, [Desc], [Expected], [Actual], [Status])
						TDHelper.AddStepToRun sStepName, sDescription,,,sStatus						
					'Case CMD_TEST 'TODO Step local logging?
					'no step reporting in local run
			End Select
			
		End If	
	End Sub	
	'==============================================================================
	' Function:		TurnOffDebug
	' Purpose:
	' Parameters:
	' Author		
	'==============================================================================
	Public Sub TurnOffDebug()

		globDebug = False
		
		AppendToFile LogFilePath, Now() & " Debug mode turned off" & vblf 		
	End Sub
	
	'==============================================================================
	' Function:		TurnOnDebug
	' Purpose:
	' Parameters:
	' Author		
	'==============================================================================
	Public Sub TurnOnDebug()
		globDebug = True			
		AppendToFile LogFilePath, Now() & " Debug mode turned on" & vblf 		
	End Sub
'==============================================================================================
' End Class clsTraceLog
'==============================================================================================
End Class

'--------------------------------------------------------------------------------------------------
' Class Start
'--------------------------------------------------------------------------------------------------
Class clsQC
	'==============================================================================================
	' CLASS VARIABLES
	'==============================================================================================

	'----------------------------------------------------------------------------------------------
	' Public Variables
	'----------------------------------------------------------------------------------------------

	'----------------------------------------------------------------------------------------------
	' Private Variables
	'----------------------------------------------------------------------------------------------

	'----------------------------------------------------------------------------------------------
	' CLASS_INITIALIZE
	'----------------------------------------------------------------------------------------------
	Private Sub Class_Initialize

	End Sub

	'==============================================================================================
	' CLASS SUBS & FUNCTIONS
	'==============================================================================================
	'==============================================================================================
	'
	'==============================================================================================
	Public Function IsQCRun()
		Dim bRetVal
		
		'TODO put real check for VAPI and QTP test 
		
		bRetVal = False		
		Select Case globRunMode
			Case QTP_TEST
				
				If globTDConnection.Connected Then
					bRetVal = True
				Else
					bRetVal = False						
					oVBSFramework.oTraceLog.Message "ALM not connected and run mode is QTP Test", LOG_ERROR
				End If	
			Case QTP_LOCAL_TEST
				bRetVal = False				
			Case VAPI_XP_TEST
				If globTDConnection.Connected Then
					bRetVal = True
				Else
					bRetVal = False	
					oVBSFramework.oTraceLog.Message "ALM not connected and run mode is VAPI Test", LOG_ERROR
				End If
			Case CMD_TEST
				bRetVal = False	
		End Select
		IsQCRun = bRetVal
	End Function
	
	'==============================================================================================
	'
	'==============================================================================================
	Public Function UploadAttachmentToQCRun(sFullPath)
		Dim oFS, oAttachFact, oAttachment

		'copy file first into temporary location if path exceeds limit (upoading to QC wouldn't work)
		If Len(sFullPath) > 254 Then
			Set oFS = CreateObject("Scripting.FileSystemObject")
			oFS.CopyFile sFullPath, PATH_RESOURCES & oFS.GetFile(sFullPath).Name
			sFullPath = PATH_RESOURCES & oFS.GetFile(sFullPath).Name
			Set oFS = Nothing
		End If
	
		Set oAttachFact = globCurrentRun.Attachments
		Set oAttachment = oAttachFact.AddItem(Null)
		oAttachment.FileName = sFullPath
		oAttachment.Type = 1
		oAttachment.Post
		oAttachment.Refresh
		
		Set oAttachment = Nothing
		Set oAttachFact = Nothing
	End Function

	'==============================================================================================
	'
	'==============================================================================================
	Public Function GetAttachmentFileFromQC(sPrefix, sExtension, sDestinationFolder)
		Dim oCurrentTest, oAttachFact, oAttachment
		Dim AttachList, sAttachList, sFilename
		Dim oFS, sAttachmentName, oShell, sSavedPath, sLoadPath, sDestinationPath

		Set oCurrentTest = globCurrentTSTest.Test 
		Set oAttachFact = oCurrentTest.Attachments
		Set AttachList = oAttachFact.NewList("")
		Set oFS = CreateObject("Scripting.FileSystemObject")
		set oShell = CreateObject("wscript.shell")

		sFilename = ""
		If Attachlist.Count > 0 Then
			For Each oAttachment In AttachList
				If UCase(Left(oAttachment.Name(1), Len(sPrefix))) = uCase(sPrefix) And UCase(Right(oAttachment.Name(1), Len(sExtension))) = UCase(sExtension) Then
					'Copy the file to the specified path
					oAttachment.Load True, ""

					sLoadPath = oFS.GetParentFolderName(oAttachment.Filename)
					sDestinationPath = sDestinationFolder
					If sDestinationFolder = "" Then sDestinationPath = sLoadPath

					'Save the original folder
					sSavedPath = oShell.CurrentDirectory
					
					'Move the file- but first get rid of it if it already exists								
					'Goto Destination Folder
					oShell.CurrentDirectory = sDestinationPath
					sAttachmentName = oAttachment.Name(1)
					If oFS.FileExists(oFS.GetFileName(sAttachmentName)) Then oFS.DeleteFile(oFS.GetFileName(sAttachmentName))

					'Goto Load folder
					oShell.CurrentDirectory = sLoadPath
					
					oFS.MoveFile oFS.GetFileName(oAttachment.Filename), sDestinationFolder & oFS.GetFileName(sAttachmentName)
					TouchFile sDestinationFolder, oFS.GetFileName(sAttachmentName)					
					sFilename = sDestinationPath & "\" & oFS.GetFileName(sAttachmentName)	'Get the file without the QC viewformat prefix

					'Revert to original folder
					oShell.CurrentDirectory = sSavedPath
					Exit For 	
				End If
			Next
		Else
			'No attachment found
		End If

		Set oShell = Nothing
		Set oFS = Nothing
		Set oCurrentTest = Nothing
		Set oAttachFact = Nothing
		Set AttachList = Nothing

		GetAttachmentFileFromQC = sFilename
	End Function
	
	'==============================================================================================
	'
	'==============================================================================================
	Public Function GetAttachmentFileFromQCByPattern(sPattern, sDestinationFolder)
		Dim oCurrentTest, oAttachFact, oAttachment
		Dim AttachList, sAttachList, sFilename
		Dim oFS, sAttachmentName, oShell, sSavedPath, sLoadPath, sDestinationPath
		Dim oRegExp, oMatch, oMatches, bMatched
		
		Const sRoutine = "clsQC.GetAttachmentFileFromQCByPattern"
		oVBSFramework.oTraceLog.HostLogMessage("> > ENTERED: " & sRoutine)
		
		Set oCurrentTest = globCurrentTSTest.Test 
		Set oAttachFact = oCurrentTest.Attachments
		Set AttachList = oAttachFact.NewList("")
		Set oFS = CreateObject("Scripting.FileSystemObject")
		set oShell = CreateObject("wscript.shell")
		
		Set oRegExp = New RegExp		' Create a regular expression.
		oRegExp.Pattern = sPattern      ' Set pattern.
		oRegExp.IgnoreCase = True       ' Set case insensitivity.
		oRegExp.Global = True         	' Set global applicability.

		sFilename = ""
		If Attachlist.Count > 0 Then
			For Each oAttachment In AttachList
				Set oMatches = oRegExp.Execute(oAttachment.Name(1))   ' Execute search.
				If oMatches.Count > 0 Then
					'Copy the file to the specified path
					oAttachment.Load True, ""

					'MsgBox "File " & oAttachment.Filename & " found matching pattern: " & sPattern
					oVBSFramework.oTraceLog.HostLogMessage "File " & oAttachment.Filename & " found matching pattern: " & sPattern
					
					sLoadPath = oFS.GetParentFolderName(oAttachment.Filename)
					sDestinationPath = sDestinationFolder
					If sDestinationFolder = "" Then sDestinationPath = sLoadPath
				
					'Save the original folder
					sSavedPath = oShell.CurrentDirectory
					
					'Move the file- but first get rid of it if it already exists								
					'Goto Destination Folder
					oShell.CurrentDirectory = sDestinationPath
					sAttachmentName = oAttachment.Name(1)
					
					On Error Resume Next
					If oFS.FileExists(oFS.GetFileName(sAttachmentName)) Then oFS.DeleteFile(oFS.GetFileName(sAttachmentName))

					If Err.Number <> 0 Then
						oVBSFramework.oTraceLog.HostLogMessage "ERROR File " & sAttachmentName & "cannot be deleted. File is probably locked."
						GetAttachmentFileFromQCByPattern = ""
						oVBSFramework.oTraceLog.HostLogMessage("< < EXITED: " & sRoutine): Exit Function
					End If
					
					On Error Goto 0
					
 					'Goto Load folder
					oShell.CurrentDirectory = sLoadPath
					
					
					oFS.MoveFile oFS.GetFileName(oAttachment.Filename), sDestinationFolder & "\" & oFS.GetFileName(sAttachmentName)
					
					TouchFile sDestinationFolder, oFS.GetFileName(sAttachmentName)	
								
					sFilename = sDestinationPath & "\" & oFS.GetFileName(sAttachmentName)	'Get the file without the QC viewformat prefix
					
					'Revert to original folder
					oShell.CurrentDirectory = sSavedPath
					Exit For 
				Else
'					MsgBox "File not found matching pattern: " & sPattern & " Current attachment is " & oAttachment.Name(1)
					oVBSFramework.oTraceLog.HostLogMessage "File not found matching pattern: " & sPattern & " Current attachment is " & oAttachment.Name(1)
				End If
				Set oMatches = Nothing
			Next
		Else
			'No attachment found
		End If

		Set oMatches = Nothing
		Set oRegExp = Nothing
		Set oShell = Nothing
		Set oFS = Nothing
		Set oCurrentTest = Nothing
		Set oAttachFact = Nothing
		Set AttachList = Nothing

		GetAttachmentFileFromQCByPattern = sFilename
		oVBSFramework.oTraceLog.HostLogMessage("< < EXITED: " & sRoutine)
	End Function

	'==============================================================================================
	'
	'==============================================================================================
	Public Function CountAttachedFilesByPattern(sPattern)
		Dim iCount
		Dim oCurrentTest, oAttachFact, oAttachment
		Dim oAttachList, sAttachList, sFilename
		Dim oRegExp, oMatch, oMatches, bMatched
		
		Const sRoutine = "clsQC.CountAttachedFilesByPattern"	
		oVBSFramework.oTraceLog.HostLogMessage("> > ENTERED : " & sRoutine)
		
		Set oCurrentTest = globCurrentTSTest.Test 
		Set oAttachFact = oCurrentTest.Attachments
		Set oAttachList = oAttachFact.NewList("")
		
		Set oRegExp = New RegExp		' Create a regular expression.
		oRegExp.Pattern = sPattern      ' Set pattern.
		oRegExp.IgnoreCase = True       ' Set case insensitivity.
		oRegExp.Global = True         	' Set global applicability.
		
		iCount = 0
		If oAttachList.Count > 0 Then
			For Each oAttachment In oAttachList
				Set oMatches = oRegExp.Execute(oAttachment.Name(1))   ' Execute search.
				If oMatches.Count > 0 Then iCount = iCount + 1
				Set oMatches = Nothing
			Next
		End If

		Set oMatches = Nothing
		Set oRegExp = Nothing
		Set oCurrentTest = Nothing
		Set oAttachFact = Nothing
		Set oAttachList = Nothing

		CountAttachedFilesByPattern = iCount
		oVBSFramework.oTraceLog.HostLogMessage("Count of attachments: " & iCount)
		oVBSFramework.oTraceLog.HostLogMessage("< < EXITED: " & sRoutine)
	End Function
	
	'==============================================================================================
	'
	'==============================================================================================
	Public Function GetResourceFileFromQC(sResourceName)
		Dim oQCConnection, oQCResourcesFactory, oFilter, oResourceList
		Dim sResourceFileName, sFileName
	
		sFileName = ""
		
		'get configuration file from QC Resources
		Set oQCConnection = globTDConnection 
		Set oQCResourcesFactory = oQCConnection.QCResourceFactory
		Set oFilter = oQCConnection.QCResourceFactory.Filter
		oFilter.Filter("RSC_NAME") = sResourceName
		Set oResourceList = oQCConnection.QCResourceFactory.NewList(oFilter.Text)
		'allow only one occurrence of resource
		If oResourceList.Count = 1 Then
			sResourceFileName = oResourceList(1).FileName
			oResourceList(1).DownloadResource PATH_RESOURCES, True
			sFileName = PATH_RESOURCES & sResourceFileName
		End If

		Set oResourceList = Nothing
		Set oFilter = Nothing
		Set oQCResourcesFactory = Nothing
		Set oQCConnection = Nothing
		
		GetResourceFileFromQC = sFileName
	End Function

	'==============================================================================================
	'
	'==============================================================================================
	Public Function AddDefect(sAssignedTo, sDetectedBy, sSummary, sDescription, sAttachmentFullPath)
		Dim oQCConnection
		Dim oBugFact, oBug
		Dim oAttachFact, oAttachment

		Set oQCConnection = globTDConnection 
		Set oBugFact = oQCConnection.BugFactory

		Set oBug = oBugFact.AddItem(Null)

		oBug.AutoPost = False
		oBug.AssignedTo = sAssignedTo
		oBug.DetectedBy = sDetectedBy
		oBug.Priority = "1-Low"
		oBug.Status = "New"
		oBug.Summary = sSummary
		oBug.Field("BG_DESCRIPTION") = sDescription
		oBug.Field("BG_DETECTION_DATE") = now()
		oBug.Field("BG_SEVERITY") = "2-Medium"
		oBug.Field("BG_REPRODUCIBLE") = "Y"
		oBug.Field("BG_PRIORITY") = "2-Medium"
		oBug.Field("BG_RESPONSIBLE") = sAssignedTo
		oBug.Field("BG_STATUS") = "New"
		'oBug.Field("BG_CATEGORY") = "Automated"
		oBug.Field("BG_PROJECT") = globTDConnection.DomainName & "." & globTDConnection.ProjectName
		oBug.Field("BG_USER_01") = "1abcd"
		oBug.Field("BG_USER_03") = "English"	'Language
		oBug.Field("BG_USER_04") = "N"			'Regression
		oBug.Field("BG_USER_05") = "Test Automation"	'Category

		oBug.Post
		oBug.Refresh
		
		'Now Attach file
		If sAttachmentFullPath <> "" then 
			Set oAttachFact = oBug.Attachments
			Set oAttachment = oAttachFact.AddItem(Null)
			oAttachment.FileName = sAttachmentFullPath
			oAttachment.Type = 1
			oAttachment.Post
			oAttachment.Refresh

			Set oAttachFact = Nothing
		End If

		Set AddDefect = oBug

	 End Function

	'==============================================================================================
	'
	'==============================================================================================
	Public Sub MailDefect(iBugID, sMailTo, sMailCC, sMailSubject, sMailComment)
		Dim i
		Dim oQCConnection
		Dim oBugFact, oBug, oBugList

		Set oQCConnection = globTDConnection 
		Set oBugFact = oQCConnection.BugFactory
		Set oBugList = oBugFact.newlist("")

		For i = oBugList.count to 1 step -1
			Set oBug = oBuglist.item(i)
			If oBuglist.item(i).ID = iBugID Then
				Exit For 
			End If
		Next

		oBug.Mail sMailTo, sMailCC, 4, sMailSubject, sMailComment

	End Sub
	'REM need to be there this function<<???
	'==============================================================================================
	' Function/Sub: TouchFile()
	' Purpose: 
	'==============================================================================================
	Private Sub TouchFile(sFolderPath, sFileName) 
		Dim oApp, oFolder, oFile
	
		Set oApp = CreateObject("Shell.Application") 
		Set oFolder = oApp.NameSpace(sFolderPath) 
		Set oFile = oFolder.ParseName(sFileName) 
		 
		oFile.ModifyDate = CStr(now) 
		
		set oFile = nothing 
		set oFolder = nothing 
		set oApp = nothing 
	End Sub 
	
	'==============================================================================================
	Public Function GetQCTestPath()
		GetQCTestPath = ""
		If IsQCRun() Then GetQCTestPath = globCurrentTSTest.TestSet.TestSetFolder.Path
	End Function
	'--------------------------------------------------------------------------------------------------
	' Class End clsQC
	'--------------------------------------------------------------------------------------------------
End class	'clsQC

'==================================================================================================
' Class Start
'==================================================================================================

Class clsDummyQRSData

	'==============================================================================================
	' CLASS VARIABLES
	'==============================================================================================

	'----------------------------------------------------------------------------------------------
	' Public Variables
	'----------------------------------------------------------------------------------------------
	'For recovery 
	Public sQRSWorkbook
	Public sQRSQCAssignee
	Public sQRSDefectDescription
	Public bQRSQCAutoDefect
	Public bQRSAutoCleanUp
	Public bQRSDefectFound
	Public sQRSQCEmailRecipients
	Public oQRSRow
	Public sQRSLog
	Public sQRSOutputParams
	Public oQRSWorkbook
	
	Public dQRSFrameworkStart

	Public QRSXL_RESULT
	Public QRSXL_KEYWORD
	Public QRSXL_COMMENT
	Public QRSXL_LOG
	Public QRSXL_REFERENCE
	Public QRSXL_OUTPUT_PARAMS
	Public QRSXL_PARM_001
	Public QRSXL_PARM_002
	Public QRSXL_PARM_003
	Public QRSXL_PARM_004
	Public QRSXL_PARM_005
	Public QRSXL_PARM_006
	Public QRSXL_PARM_007
	Public QRSXL_PARM_008
	Public QRSXL_PARM_009
	Public QRSXL_PARM_010

	'----------------------------------------------------------------------------------------------
	' Private Variables
	'----------------------------------------------------------------------------------------------
	'private p_sLog

	'==============================================================================================
	' CLASS PROPERTIES
	'==============================================================================================

	'Public Property Let sQRSLog(sString)
	'	p_sLog = sString
	'End Property	

	'Public Property Get sQRSLog
	'	sQRSLog = p_sLog
	'End Property	
	'==============================================================================================
	' CLASS SUBS & FUNCTIONS
	'==============================================================================================

	'----------------------------------------------------------------------------------------------
	' CLASS_INITIALIZE
	'----------------------------------------------------------------------------------------------
	Private Sub Class_Initialize
		
	End Sub

	'----------------------------------------------------------------------------------------------
	' CLASS_TERMINATE
	'----------------------------------------------------------------------------------------------
	Private Sub Class_Terminate

	End Sub

'==============================================================================================
' End Class clsDummyQRSData
'==============================================================================================
End Class

'==================================================================================================
' Class Start
'==================================================================================================
Class clsModules

	'==============================================================================================
	' CLASS VARIABLES
	'==============================================================================================

	'----------------------------------------------------------------------------------------------
	' Public Variables
	'----------------------------------------------------------------------------------------------

	'----------------------------------------------------------------------------------------------
	' Private Variables
	'----------------------------------------------------------------------------------------------
	Private dictModuleSourceFile
	Private dictModuleDependencies
	Private dictModuleLoaded
	Private dictCoreModule
	Private dictTOAliases 
	Private oFso

	'==============================================================================================
	' CLASS PROPERTIES
	'==============================================================================================

	'==============================================================================================
	' CLASS SUBS & FUNCTIONS
	'==============================================================================================

	'----------------------------------------------------------------------------------------------
	' CLASS_INITIALIZE
	'----------------------------------------------------------------------------------------------
	Private Sub Class_Initialize
		Set dictModuleSourceFile = CreateObject("Scripting.Dictionary")
		Set dictModuleDependencies = CreateObject("Scripting.Dictionary")
		Set dictModuleLoaded = CreateObject("Scripting.Dictionary")
		Set dictCoreModule = CreateObject("Scripting.Dictionary")
		Set dictTOAliases = CreateObject("Scripting.Dictionary")
		Set oFso = CreateObject("Scripting.FileSystemObject")
	End Sub

	'----------------------------------------------------------------------------------------------
	' CLASS_TERMINATE
	'----------------------------------------------------------------------------------------------
	Private Sub Class_Terminate
		Set dictModuleSourceFile = Nothing
		Set dictModuleDependencies = Nothing
		Set dictModuleLoaded = Nothing
		Set dictTOAliases = Nothing
		Set dictCoreModule  = Nothing
		Set oFso = Nothing
	End Sub

	'==============================================================================
	' Function:		Load
	' Returns:		True if the module is loaded this time 
	'				False if already loaded
	'==============================================================================
	Public Function Load(sModule)
		Dim bRetVal
		Dim arrDependencies, sDependency
		Dim arrAliases, sAliases
		Dim sSourceFile, sReturnModule
			
		bRetval = True
		sAliases = GetAliases(UCase(sModule))
		If globDebug Then
			oVBSFramework.oTraceLog.Message "Loading " & sModule & " in oTraceLog.Load", LOG_DEBUG
		End If
		If sAliases = "" Then				
			'sModule may be name of alias
		
			sReturnModule = GetNameOfTestObject(sModule)						
			If sReturnModule <> "" Then			 
				'sModule was alias of Test Object, if not sReturnMode is same as sModule was				
				sModule = sReturnModule	
			End If	
			' if was not alias then it does not have aliases, or such test object not exist, chcek for it is done below		
		End If		
		
		'Check if it's loaded
		If Not IsLoaded(UCase(sModule)) Then		
		
			'Do it's dependencies first
			arrDependencies = Split(GetDependencies(UCase(sModule)), ",")

			For each sDependency in arrDependencies			'Repeat the load process for each dependency
				If StrComp(UCase(sDependency),"QTP", 1) = 0 Then ' checking QTP dependancy
					If (Not globRunMode = QTP_TEST) AND (Not globRunMode = QTP_LOCAL_TEST) Then ' running no QTP Test so loading will fail
						Load = False
						oVBSFramework.oTraceLog.Message "No-QTP test is running with QTP classes. Module " & sModule & "is QTP dependant and No-QTP test is running. Consider to change run mode",LOG_ERROR
						Exit Function
					End if
				Else				
					If Load(UCase(sDependency)) = False Then	'Load failed ...
						Load = False	
						Exit Function
					End If
				End If
			Next
	
			'Load the source file using LoadModule.  This checks for the existence of the file.
			sSourceFile = GetSourceFile(UCase(sModule))
			
			If Not sSourceFile = "" Then
				If oVBSFramework.LoadModule(sSourceFile) = True Then
					'The file should set itself to loaded but we can make sure by doing it here anyway
					SetLoaded UCase(sModule)
					bRetval = True
				Else
					bRetVal = False
				End If
			Else 
				bRetVal = False
			End If
		End If
	
		Load = bRetVal
	End Function

	'==============================================================================
	' Function:		SetLoaded
	' Returns:		True if the module is loaded this time 
	'				False if already loaded
	'==============================================================================
	Public Function SetLoaded(sModule)
		Dim bRetVal
	
		bRetval = False
		If Not IsLoaded(UCase(sModule)) Then
			dictModuleLoaded.Add UCase(sModule), True
			bRetval = True
		End If
		SetLoaded = bRetVal
	End Function

	'==============================================================================
	' Function:		IsLoaded
	'==============================================================================
	Public function IsLoaded(sModule)
		Dim bRetVal

		bRetVal = False
		if dictModuleLoaded.exists(UCase(sModule)) then bRetVal = True	
		IsLoaded = bRetVal
	End function

	'==============================================================================
	' Function:		SetDependencies
	' Returns:		True if the dependencies are set this time
	'				False if already set
	'==============================================================================
	Public Function SetDependencies(sModule, sDependencies)
		Dim bRetVal

		bRetval = False
		If Not dictModuleDependencies.Exists(UCase(sModule)) Then
			dictModuleDependencies.Add UCase(sModule), sDependencies
			bRetval = True
		End If

		SetDependencies = bRetVal
	End Function
	'==============================================================================
	' Function:		SetAliases
	' Returns:		True if the aliases are set this time
	'				False if already set
	' Author		
	'==============================================================================
	Public Function SetAliases(sTestObject, sAlias)
		Dim bRetVal

		bRetval = False
		If Not dictTOAliases.Exists(UCase(sTestObject)) Then		
			dictTOAliases.Add UCase(sTestObject), sAlias
			bRetval = True
		End If
		SetAliases = bRetVal
	End function
	'==============================================================================
	' Function:		GetAliases
	' Returns:		Aliases or Empty string if no aliases
	' Author		
	'==============================================================================
	Public Function GetAliases(sTestObject)
		Dim sRetVal
		sRetVal = ""	
		If dictTOAliases.Exists(UCase(sTestObject)) Then			
			sRetVal = dictTOAliases(UCase(sTestObject))					
		End If
		GetAliases = sRetVal
	End Function
	'==============================================================================
	' Function:		GetNameOfTestObject
	' Returns:		Name of Test Object or Empty string if no aliases
	' Author		
	'==============================================================================
	Public Function GetNameOfTestObject(sAlias)
		Dim sRetVal
		Dim oKeys, sKey
		Dim arrAliases
		
		sRetVal = ""
		
		If dictTOAliases.Exists(UCase(sAlias)) Then
			' if dictionary contains value Alias is name of Test Object
			GetNameOfTestObject = sAlias
			Exit Function
		Else		
			oKeys = dictTOAliases.Keys 
			' only simple loop througth aliases, find first occurrence of alias and return name of Test Object, 
			' we suppose that Alias is unique
			For Each sKey in oKeys 
    			arrAliases = Split(GetAliases(UCase(sKey)), ",")    			
				' check if array contains alias
				If (UBound(Filter(arrAliases, sAlias, True, vbTextCompare)) > -1) Then
					sRetVal = sKey				
					Exit For
				End If
			Next 
			
			'sAlias is TO with no aliases or we do not know TO
			If sRetVal = "" Then
				If GetSourceFile(UCase(sAlias)) <> "" Then
					sRetVal = sAlias
				Else
					'we do not have such test object in sources
					sRetVal = ""
					oVBSFramework.oTraceLog.Message "GetNameOfTestObject: No Test object for" & sAlias & " recognised", LOG_ERROR
				End If
			End If	
		End If
	

		GetNameOfTestObject = sRetVal
	End function

	'==============================================================================
	' Function:		GetDependencies
	' Returns:		Dependencies or Empty string if no dependencies
	'==============================================================================
	Public Function GetDependencies(sModule)
		Dim sRetVal

		sRetVal = ""
		If dictModuleDependencies.Exists(UCase(sModule)) Then
			sRetVal = dictModuleDependencies(UCase(sModule))
		End If

		GetDependencies = sRetVal
	End function

	'==============================================================================
	' Function:		SetSourceFile
	' Returns:		True if the sSourceFile set this time
	'				False if already set
	'==============================================================================
	Public Function SetSourceFile(sModule, sSourceFile)
		Dim bRetVal	
		bRetval = False
		If Not dictModuleSourceFile.Exists(UCase(sModule)) Then
			dictModuleSourceFile.Add UCase(sModule), UCase(sSourceFile)
			If Not dictModuleSourceFile.Exists(UCase(sSourceFile)) Then
				dictModuleSourceFile.Add UCase(sSourceFile), UCase(sSourceFile)
			End If
			bRetval = True
		End If

		SetSourceFile = bRetVal
	End function

	'==============================================================================
	' Function:		GetSourceFile
	' Returns:		Source file or Empty string if no source file
	'==============================================================================
	Public Function GetSourceFile(sModule)
		Dim sRetVal

		sRetVal = ""
		If dictModuleSourceFile.Exists(UCase(sModule)) Then
			sRetVal = dictModuleSourceFile(UCase(sModule))
		End If

		GetSourceFile = sRetVal
	End function

	'==============================================================================
	' Function:		SetCoreModule
	' Returns:		True if the sModule set this time
	'				False if already set
	'==============================================================================
	Public Function SetCoreModule(sModule)
		Dim bRetVal

		bRetval = False
		If Not dictCoreModule.Exists(UCase(sModule)) Then
			dictCoreModule.Add UCase(sModule), UCase(sModule)
			bRetval = True
		End If

		SetCoreModule = bRetVal
	End function

	'==============================================================================
	' Function:		SourceFilesAndDependenciesFromINI 
	' Overview 		Loads information from .INI file
	' Returns:		True if ini file found, False if ini file not found
	' Author		
	'==============================================================================
	Public Function SourceFilesAndDependenciesFromINI(sINIFile, sOutput)
		 Dim oFSO, oIniFile
		 Dim bRetVal
		 Dim arrDependencies, arrSourceFiles, arrCoreModules, arrINILine, arrAliases
		 Dim sLine
		 bRetval = True
		 Set oFSO = CreateObject("Scripting.FileSystemObject")		 
		 If oFSO.FileExists(sINIFile) Then
			oVBSFramework.oTraceLog.Message "Loading .ini file", LOG_MESSAGE
	        Set oIniFile = oFSO.OpenTextFile(sINIFile, 1, False) ' opens .ini file
			
			Do While oIniFile.AtEndOfStream = False
				
				sLine = Trim(oIniFile.ReadLine)
				' loads source files info
				If LCase(sLine) = "[" & LCase("SourceFiles") & "]" Then
					sLine = Trim(oIniFile.ReadLine)	
					If globDebug Then
							oVBSFramework.oTraceLog.Message "Loading source files", LOG_DEBUG
					End If		
					Do While Len(sLine) > 0 AND Left(sLine, 1) <> "["					
						 arrINILine = Split(sLine,"=")
						 arrSourceFiles = Split(arrINILine(1),"~")							 
						 SetSourceFile UCase(arrSourceFiles(0)), arrSourceFiles(1)					 
					     sOutput = sOutput & " SetSourceFile: " & UCase(arrSourceFiles(0)) & "=" & arrSourceFiles(1) & vblf
						 sLine = Trim(oIniFile.ReadLine)
					
					 Loop
				End If
				' load dependencies info
				If LCase(sLine) = "[" & LCase("Dependencies") & "]" Then
					sLine = Trim( oIniFile.ReadLine )
					If globDebug Then
							oVBSFramework.oTraceLog.Message "Loading dependencies", LOG_DEBUG
					End If			
					 Do While Len(sLine) > 0 AND Left(sLine, 1) <> "["					
						 arrINILine = Split(sLine,"=")
						 arrDependencies = Split(arrINILine(1),"~")	
						 SetDependencies UCase(arrDependencies(0)), arrDependencies(1)					 	 
					 	 sOutput = sOutput & " SetDependency: " & UCase(arrDependencies(0)) & "=" & arrDependencies(1) & vblf			
					     sLine = Trim(oIniFile.ReadLine)
						
					 Loop
				End If
				' loads core modles info
				If LCase(sLine) = "[" & LCase("CoreModules") & "]" Then
					sLine = Trim( oIniFile.ReadLine )
					If globDebug Then
							oVBSFramework.oTraceLog.Message "Loading core modules", LOG_DEBUG
					End If				
					 Do While Len(sLine) > 0 AND Left( sLine, 1 ) <> "["					
						 arrINILine = Split(sLine,"=")							 				 
						 SetCoreModule UCase(arrINILine(1))
						 sOutput = sOutput & " SetCoreMOdule: " & UCase(arrINILine(1)) & vblf
						 If oIniFile.AtEndOfStream Then
						 	Exit Do
						 Else
						 	sLine = Trim(oIniFile.ReadLine)
						 End If	
						 
					 Loop
				End If
				' loads aliases info
				If LCase(sLine) = "[" & LCase("Aliases") & "]" Then
					sLine = Trim(oIniFile.ReadLine)	
					If globDebug Then
							oVBSFramework.oTraceLog.Message "Loading aliases", LOG_DEBUG
					End If
					Do While Len(sLine) > 0 AND Left(sLine, 1) <> "["					
						 arrINILine = Split(sLine,"=")						  
						 arrAliases = Split(arrINILine(1),"~")	
						 SetAliases UCase(arrAliases(0)), arrAliases(1)					 	 '
					 	 sOutput = sOutput & "SetAliases: " & UCase(arrAliases(0)) & "=" & arrAliases(1) & vblf	
					     
					     If oIniFile.AtEndOfStream Then
						 	Exit Do
						 Else
						 	sLine = Trim(oIniFile.ReadLine)
						 End If	
						
					 Loop
				End If
					
			Loop
			
			If globDebug Then		
				oVBSFramework.oTraceLog.Message sOutput, LOG_DEBUG
			End If	
			
		 Else
			 bRetval = False
			 sOuput = sOutput & "INI file " & sINIFile & " not found." & vblf
			 oVBSFramework.oTraceLog.Message sOutput, LOG_ERROR
		 End If	
		 

		 SourceFilesAndDependenciesFromINI = bRetVal
	End Function
	'==============================================================================================
	' Subroutine:	LoadCoreModules
	'==============================================================================================
	Public Sub LoadCoreModules()
		Dim i, arrCoreModules

		arrCoreModules = dictCoreModule.Keys	
		oVBSFramework.oTraceLog.Message "Loading Core Modules", LOG_MESSAGE
		
		For i = 0 To UBound(arrCoreModules) 		
			Load arrCoreModules(i)
		Next
		
	End Sub
	

'==============================================================================================
' End Class clsModules
'==============================================================================================
End Class

'==============================================================================================
' Function:		IncludeFiles(sFilename, sPath)
'				Looks for lines of the form: #INCLUDE folder\incfile
'				Replaces the INCLUDE lines with the contents of incfile
' Returns:		The original filename if no #includes or any error
'				The new file name with the includes expanded (eg sFilename.all.vbs)
'==============================================================================================
Public Function IncludeFiles(sFilename, sPath)
	Dim sRetVal
	Dim sFileText
	Dim sIncludeFile, sIncludeFileText
	Dim oFS, sFullFilename
	Dim iNextIncludePos 
	Dim sIncludeKeyword

	Const ForReading = 1
	
	sIncludeKeyword = vbnewline & "#INCLUDE "
	sRetVal = sFilename
	
	'Check that the module file actually exists.
	'Get the full filename and path either a "\\" or a "D:\" indicates full filename and path
	sFullFilename = GetFullFilename(sFilename, sPath)
	Set oFS = CreateObject("Scripting.FileSystemObject")
	If oFS.FileExists(sFullFilename) Then
		sFileText = ReadAllFileText(GetFullFilename(sFilename, sPath))
	
		iNextIncludePos = InStr(1, sFileText, sIncludeKeyword, 1)
	
		If iNextIncludePos > 0 Then
			Do While (iNextIncludePos > 0)
				sIncludeFile = GetStringFromText(sFileText, Len(sIncludeKeyword)+iNextIncludePos, vbNewline)
				sFullFilename = GetFullFilename(Trim(sIncludeFile), sPath)
				If oFS.FileExists(sFullFilename) Then
					sIncludeFileText = ReadAllFileText(sFullFilename)
					sIncludeFileText = vbnewline & "'BEGIN #INCLUDE " & sFullFilename & " *************************************** BEGIN" & vbnewline &  sIncludeFileText & _
									   vbnewline & "'END   #INCLUDE " & sFullFilename & " ***************************************** END" & vbnewline
					sFileText = Replace(sFileText, sIncludeKeyword & sIncludeFile, sIncludeFileText, 1, 1, 1)
					iNextIncludePos = InStr(1, sFileText, sIncludeKeyword, 1)
				Else
					'Include file not found. Generate an error by 
					IncludeFiles = sFilename & ".Error_" & Replace(Trim(sIncludeKeyword & sIncludeFile)," ","_") 
					Exit Function
				End If 
			Loop
			
			'Write the text to a new file
			sFullFilename = GetFullFilename(sFilename, sPath) & ".all.vbs"
			sRetVal = WriteAllFileText(sFullFilename, vbnewline & "'BEGIN #INCLUDE " & sFullFilename & vbnewline & vbnewline & sFileText &vbnewline & "'END #INCLUDE " & sFullFilename & vbnewline & vbnewline)
		Else
			'No includes
		End If
	Else
		'File does not exist
	End If
	
	Set oFS = Nothing

	IncludeFiles = sRetVal
End Function

'==============================================================================================
' Function:		GetFileFormat(sFile)
'==============================================================================================
Public Function GetFileFormat(sFile)
	Dim oFS, oFile
	Dim iFormat, intAsc1Chr, intAsc2Chr
	Const ForReading = 1
	Const iUnicode = -1
	Const iAscii = 0
	
    iFormat = iAscii	'Assume ASCII
	Set oFS = CreateObject("Scripting.FileSystemObject")
	If oFS.FileExists(sFile) Then
		'Detect unicode file
		Set oFile = oFS.OpenTextFile(sFile, ForReading, False)
		intAsc1Chr = Asc(oFile.Read(1))
		intAsc2Chr = Asc(oFile.Read(1))
		oFile.Close
		Set oFile = Nothing
		If intAsc1Chr = 255 And intAsc2Chr = 254 Then 
		    iFormat = iUnicode
		End If
	Else
		'TODO: file not found
	End If

	Set oFS = Nothing 
	
	GetFileFormat = iFormat
End Function

'==============================================================================================
' Function:		ReadAllFileText(sFile)
'==============================================================================================
Public Function ReadAllFileText(sFile)
	Dim oFS, oFile, sFileText
	Dim iFormat
	Const ForReading = 1
	
	Set oFS = CreateObject("Scripting.FileSystemObject")
	If oFS.FileExists(sFile) Then
		iFormat = GetFileFormat(sFile)

		'Get file content
		Set oFile = oFS.OpenTextFile(sFile, ForReading, False, iFormat)
		sFileText = oFile.ReadAll
		oFile.Close
		Set oFile = Nothing 
'MsgBox "Stopped" : STOP
	End If

	Set oFS = Nothing 
	
	ReadAllFileText = sFileText
End Function

'==============================================================================================
' Function:		WriteAllFileText(sFile, sText)
'==============================================================================================
Public Function WriteAllFileText(sFile, sText)
	Dim oFS, oFile
	Dim iFormat
	Const ForWriting = 2
	
	WriteAllFileText = sFile
	Set oFS = CreateObject("Scripting.FileSystemObject")
	If oFS.FolderExists(oFS.GetParentFolderName(sFile)) Then
		iFormat = GetFileFormat(sFile)
		Set oFile = oFS.OpenTextFile(sFile, ForWriting, True, iFormat)
		oFile.Write sText
		oFile.Close
		Set oFile = Nothing 
	else
		WriteAllFileText = sFile & ".Error_PathNotFound" 
	End If

	Set oFS = Nothing 
	
End function

'==============================================================================================
' Function:		GetFullFilename(sFilename, sPath)
'==============================================================================================
Public Function GetFullFilename(sFilename, sPath)
	Dim sFullFileName
	
	'Get the full filename and path either a "\\" or a "D:\" indicates full filename and path
	If Left(Trim(sFilename),2) = "\\" OR Mid(Trim(sFilename),2,1) = ":" Then
		sFullFilename = sFilename
	Else
		sFullFilename = sPath & "\" & sFilename
	End If 

	GetFullFilename = sFullFileName
End Function

'==============================================================================================
' Function:		GetTextFromString(sText, iStartPos, sEndChar)
'				returns string from from a start position in a text block upto a defined character (eg vbnewline)
' Returns:		Text found
'==============================================================================================
Public Function GetStringFromText(sText, iStartPos, sEndChar)
	Dim iEndPos, sString
	
	iEndPos = InStr(iStartpos, sText, sEndChar, 1)
	sString = Mid(sText, iStartPos, iEndPos - iStartPos)		
	GetStringFromText = sString
End Function

'==============================================================================================
' Function/Sub: AppendToFile()
' Purpose: 
'==============================================================================================
Public Sub AppendToFile(sFile, sText)
	Dim oFS, oFile
	Dim bDoThis
	
	'bDoThis = False
	bDoThis = True 'logging is allowed
	If bDoThis then
		Set oFS = CreateObject("Scripting.FileSystemObject")
		On Error Resume Next
		oFS.CreateFolder(oFS.GetParentFolderName(sFile))
		On Error GoTo 0
		Set oFile = oFS.OpenTextFile(sFile, 8, true)	'appending
		oFile.WriteLine sText
		oFile.Close
		Set oFile = Nothing
		Set oFS = Nothing
	End If
	
End Sub
'==============================================================================================
' Function/Sub: ExecuteFileToGlobal(sFilePath)
'				Load .vbs file to global scope by ExecuteGlobal, instead QTP specific ExecuteFile	
'==============================================================================================
Public Function ExecuteFileToGlobal(sFilePath)
	Dim oFS, oFile, sScript, intAsc1Chr, intAsc2Chr, bOpenAsUnicode
	Const ForReading = 1
	On Error Resume Next
	Set oFS = CreateObject("Scripting.FileSystemObject")

	If oFS.FileExists(sFilePath) Then
		'Detect unicode file
		Set oFile = oFS.OpenTextFile(sFilePath, ForReading, False)
		intAsc1Chr = Asc(oFile.Read(1))
		intAsc2Chr = Asc(oFile.Read(1))
		oFile.Close
		If intAsc1Chr = 255 And intAsc2Chr = 254 Then 
		    bOpenAsUnicode = True
		Else
		    bOpenAsUnicode = False
		End If
		
		'Get script content
		Set oFile = oFS.OpenTextFile(sFilePath, ForReading, False, bOpenAsUnicode)
		sScript = oFile.ReadAll()
		oFile.Close
	
		Set oFile = Nothing 

		ExecuteGlobal sScript
		
		If Err.Number <> 0 Then
   			 'TDOutput.Print "Run-time error [" & Err.Number & "] : " & Err.Description   	
   			 'MsgBox Err.Description
   			 oVBSFramework.oTraceLog.Message "Error in ExecuteFileToGlobal. Run-time error [" & Err.Number & "] : " & Err.Description & ", Source: " & Err.Source, LOG_ERROR  			 
   			 ExecuteFileToGlobal = False 
   			 Exit Function
  		End If
	
	Else
		'MsgBox "File in path: " & sFilePath & " does not found"
		oVBSFramework.oTraceLog.Message "Error in ExecuteFileToGlobal. File in path: " & sFilePath & " does not found", LOG_ERROR
		ExecuteFileToGlobal = False
		Exit Function
	End If
	ExecuteFileToGlobal = True
	Set oFS = Nothing 	 
End Function

'==============================================================================================
Public Function YYYYMMDDHHMM(dDate)
	YYYYMMDDHHMM = Year(dDate) & Right("0" & Month(dDate), 2) & Right("0" & Day(dDate), 2) _
					& Right("0" & Hour(dDate), 2) & Right("0" & Minute(dDate), 2)
End Function

'==============================================================================================
Public Function YYYYMMDDHHMMSS(dDate)
	YYYYMMDDHHMMSS = Year(dDate) & Right("0" & Month(dDate), 2) & Right("0" & Day(dDate), 2) _
					& Right("0" & Hour(dDate), 2) & Right("0" & Minute(dDate), 2) & Right("0" & Second(dDate), 2)
End Function