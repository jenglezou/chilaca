'==============================================================================================
' Function:		
'==============================================================================================
Function RecoveryFunction(Object, Method, Arguments, retVal)
	Dim sDefectDescription		'Use this to get data back from QTPFWCleanUp

	sDefectDescription = oQRSData.sQRSDefectDescription
	'msgbox "In RecoveryFunction"
	
	'TODO: proper handling of various scenarios
	If Not IsEmpty(oQRSData.oQRSRow) Then
		If Not oQRSData.oQRSRow Is Nothing Then
			QTPFWCleanUp oQRSData.oQRSRow, oQRSData.sQRSLog, oQRSData.sQRSOutputParams, oQRSData.oQRSWorkbook, sDefectDescription
			QCCleanUp oQRSData.sQRSWorkbook, oQRSData.bQRSQCAutoDefect, oQRSData.bQRSDefectFound, oQRSData.sQRSQCAssignee, sDefectDescription, oQRSData.sQRSQCEmailRecipients
			If oQRSData.bQRSAutoCleanUp Then
				QRSWindowsProcessesCleanUp oQRSData.dQRSFrameworkStart
			End If
		End If
	End If
End Function

'==============================================================================================
' Function:
'==============================================================================================
Sub QRSWindowsProcessesCleanUp(dFrameworkStart)
	Dim oWMIDateTime, oWMIService, colProcessList, oProcess, iCount, sUserName, sDomain, sScreenshotFullPath, dictWhiteList
	
	Set dictWhiteList = CreateObject("Scripting.Dictionary")
	dictWhiteList.Add "wmiprvse.exe", "X"			' MS Windows component
	dictWhiteList.Add "rvd.exe", "X"				' TIBCO
	dictWhiteList.Add "QTAutomationAgent.exe", "X"	' QTP Automat
		
	'construct datetime in WMI format
	Set oWMIDateTime = CreateObject("WbemScripting.SWbemDateTime")
	oWMIDateTime.SetVarDate(dFrameworkStart)
	
	Set oWMIService = GetObject("winmgmts:{impersonationLevel=impersonate}!\\" & "." & "\root\cimv2")
	Set colProcessList = oWMIService.ExecQuery("Select * from Win32_Process WHERE CreationDate > '" & oWMIDateTime & "'")
	
	On Error Resume Next
	iCount = colProcessList.Count
	If (err.number <> 0) Then
		Reporter.ReportEvent micWarning, "WindowsProcessesCleanUp()", "WMI can be used on remote session only with /console parameter"
	Else
		For Each oProcess in colProcessList
			If Not dictWhiteList.Exists(oProcess.Name) Then
				oProcess.GetOwner sUserName, sDomain
				Reporter.ReportEvent micDone, "WindowsProcessesCleanUp()", "Terminating process [" & oProcess.Name & "] owned by " & sDomain & "\" & sUserName
				oProcess.Terminate()
			End If
		Next
	End If
	On Error Goto 0
	
	'screenshot
	sScreenshotFullPath = "X:\Temp\VBSFramework\Screenshots\screen" & YYYYMMDDHHMMSS(Now()) & ".png"
	Desktop.CaptureBitmap sScreenshotFullPath
	Reporter.ReportEvent micDone, "CaptureScreen", sScreenshotFullPath, sScreenshotFullPath
	
	Set dictWhiteList = Nothing
	Set colProcessList = Nothing
	Set oWMIService = Nothing
	Set oWMIDateTime = Nothing
End Sub

'==============================================================================================
' Function:		
'==============================================================================================
Sub QTPFWCleanUp(oRow, sLog, sOutputParams, oWorkbook, sDefectDescription)
	oRow.Cells(1, oQRSData.QRSXL_RESULT).Value = "fail"
	oRow.Cells(1, oQRSData.QRSXL_RESULT).Interior.ColorIndex = 3			' Red
	Reporter.ReportEvent micFail, "Unexpected error: RecoveryFunction() called", "Unexpected error: RecoveryFunction() called"

	'Write the log entry
	sLog = sLog & "E Unexpected error. See the QTP test results for more details. Hint - look at the end." & vbLF
	QRSWriteIntoLogAndOutputParamsCells oRow, sLog, sOutputParams
	
	'Get this information before closing the workbook
	sDefectDescription = sDefectDescription & "Worksheet:" & oRow.parent.name & vbnewline
	sDefectDescription = sDefectDescription & "Row:" & oRow.Row & vbnewLine & RowToText(oRow)

	' Save workbook
	oWorkbook.Save

	' Close workbook
	oWorkbook.Close
End Sub

'==============================================================================================
' Function:		
'==============================================================================================
Public Sub QRSWriteIntoLogAndOutputParamsCells(oRow, sLog, sOutputParams)
	If sLog <> "" Then
		If oRow.Cells(1, oQRSData.QRSXL_LOG).Value = "" Then
			oRow.Cells(1, oQRSData.QRSXL_LOG).Value = sLog
		Else
			oRow.Cells(1, oQRSData.QRSXL_LOG).Value = oRow.Cells(1, oQRSData.QRSXL_LOG).Value & vbLf & sLog
		End If
	End If
	If sOutputParams <> "" Then
		If oRow.Cells(1, oQRSData.QRSXL_OUTPUT_PARAMS).Value = "" Then
			oRow.Cells(1, oQRSData.QRSXL_OUTPUT_PARAMS).Value = sOutputParams
		ElseIf Right(oRow.Cells(1, oQRSData.QRSXL_OUTPUT_PARAMS).Value, 1) = "," Then
			oRow.Cells(1, oQRSData.QRSXL_OUTPUT_PARAMS).Value = oRow.Cells(1, oQRSData.QRSXL_OUTPUT_PARAMS).Value & sOutputParams
		Else
			oRow.Cells(1, oQRSData.QRSXL_OUTPUT_PARAMS).Value = oRow.Cells(1, oQRSData.QRSXL_OUTPUT_PARAMS).Value & "," & sOutputParams
		End If
	End If
End Sub

'==============================================================================================
' Function:		
'==============================================================================================
Sub QCCleanUp(sWorkbook, bQCAutoDefect, bDefectFound, sQCAssignee, sDefectDescription, sQCEmailRecipients)
	Dim oBug, oFS

	Set oFS = CreateObject("Scripting.FileSystemObject")

	'if test is run from QC, then upload results as attachment to QC Current Run
	If QRSIsQCRun() then 
		QRSUploadAttachmentToQCRun(sWorkbook)
		if bQCAutoDefect then 'AND bDefectFound then
			'Create defect
			Set oBug = QRSAddDefect(sQCAssignee, QCUtil.QCConnection.UserName, "Automated Run Workbook: " & oFS.GetFileName(sWorkbook), sDefectDescription, sWorkbook)
			'Link the defect to the current run
			QCUtil.CurrentRun.BugLinkFactory.AddItem(oBug)
			'Send mail 
			QRSMailDefect oBug.ID, sQCEmailRecipients, "", _
						"QC Defect - Domain:" & QCUtil.QCConnection.DomainName & ", Project:" & QCUtil.QCConnection.ProjectName & ", Automated Run Workbook: " & oFS.GetFileName(sWorkbook), _
						"See the description below for details of sheets and rows."
		End If
	End If
	
	Set oFS = Nothing
End Sub

'==============================================================================================
'
'==============================================================================================
Function QRSIsQCRun()
	Dim bRetVal

	bRetVal = False

	If QCUtil.IsConnected then 
		If Not QCUtil.CurrentRun Is Nothing Then
			bRetVal = True
		End If
	End If

	QRSIsQCRun = bRetVal
End Function

'==============================================================================================
'
'==============================================================================================
Function QRSUploadAttachmentToQCRun(sFullPath)
	Dim oAttachFact, oAttachment

	Set oAttachFact = QCUtil.CurrentRun.Attachments
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
Public Function QRSAddDefect(sAssignedTo, sDetectedBy, sSummary, sDescription, sAttachmentFullPath)
	Dim oQCConnection
	Dim oBugFact, oBug
	Dim oAttachFact, oAttachment

	Set oQCConnection = QCUtil.QCConnection 
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
	oBug.Field("BG_PROJECT") = QCUtil.QCConnection.DomainName & "." & QCUtil.QCConnection.ProjectName
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

	'Set oQCConnection = Nothing
	'Set oBugFact = Nothing
	
	Set QRSAddDefect = oBug

 End Function

'==============================================================================================
'
'==============================================================================================
Sub QRSMailDefect(iBugID, sMailTo, sMailCC, sMailSubject, sMailComment)
	Dim i
	Dim oQCConnection
	Dim oBugFact, oBug, oBugList

	Set oQCConnection = QCUtil.QCConnection 
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

'==============================================================================================
'
'==============================================================================================
Function RowToText(oRow)
	Dim i
	Dim sRetVal

	sRetVal = "result=" & oRow.Cells(1, oQRSData.QRSXL_RESULT).Text & vbnewline & _
				"keyword=" & oRow.Cells(1, oQRSData.QRSXL_KEYWORD).Text & vbnewline & _
				"comment=" & oRow.Cells(1, oQRSData.QRSXL_COMMENT).Text & vbnewline & _
				"output=" & oRow.Cells(1, oQRSData.QRSXL_LOG).Text & vbnewline

	For i = oQRSData.QRSXL_PARM_001 to oQRSData.QRSXL_PARM_010
		If oRow.Cells(1, i).Text = "" Then Exit For
		sRetVal = sRetVal & "parm_" & Left("00" & i-oQRSData.QRSXL_PARM_001+1, 3) & "=" & oRow.Cells(1, i).Text & vbnewline
	Next

	sRetVal = sRetVal & "message=" & oRow.Cells(1, oQRSData.QRSXL_OUTPUT_PARAMS).Text & vbnewline
	RowToText = sRetVal
End Function

'==============================================================================================
Function YYYYMMDDHHMMSS(dDate)
	YYYYMMDDHHMMSS = Year(dDate) & Right("0" & Month(dDate), 2) & Right("0" & Day(dDate), 2) _
					& Right("0" & Hour(dDate), 2) & Right("0" & Minute(dDate), 2) & Right("0" & Second(dDate), 2)
End Function

'==================================================================================================
' Class Start
'==================================================================================================

Class clsQRSData

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
		
	End Sub

	'----------------------------------------------------------------------------------------------
	' CLASS_TERMINATE
	'----------------------------------------------------------------------------------------------
	Private Sub Class_Terminate

	End Sub

'==============================================================================================
' End Class clsQRSData
'==============================================================================================
End Class

Public oQRSData
Set oQRSData = New clsQRSData
