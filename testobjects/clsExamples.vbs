Option Explicit
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
if not oVBSFramework.oTestObjects.isLoaded("EXAMPLES") then
	Public oExamples										
	Set oExamples = New clsExamples		
	oVBSFramework.oTestObjects.Add "EXAMPLES", oExamples
End If
