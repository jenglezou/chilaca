Option Explicit
'==================================================================================================
' clsFidessa.vbs
'
' Purpose
'==================================================================================================
'--------------------------------------------------------------------------------------------------
' Constants
'--------------------------------------------------------------------------------------------------

'--------------------------------------------------------------------------------------------------
' Class Start
'--------------------------------------------------------------------------------------------------

Class clsFidessa

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

	End Sub

	'==============================================================================================
	' Function/Sub:
	'==============================================================================================
	Public Function RowDispatch(oRow, sLog, sOutputParams)
		Dim iRetVal, sKeyword

		Const sRoutine = "clsFidessa.RowDispatch"
		oVBSFramework.oTraceLog.Entered(sRoutine & ", Keyword=""" & CStr(oRow.Cells(1, XL_KEYWORD).Value) & """")

		iRetVal = XL_DISPATCH_PASS
		
		' Get keyword from spreadsheet as string
		sKeyword = CStr(oRow.Cells(1, XL_KEYWORD).Value)
		sKeyword = right(sKeyword, len(sKeyword)-inStrRev(sKeyword, "."))

		Select Case uCase(sKeyword)
			Case "START"			
				iRetVal = FidessaStartAction(oRow, sLog, sOutputParams)
			Case "LOGIN"			
				iRetVal = FidessaLoginAction(oRow, sLog, sOutputParams)
			Case "NEWORDER"			
				iRetVal = FidessaNewOrderAction(oRow, sLog, sOutputParams)
			Case "ROUTEORDER"			
				iRetVal = FidessaRouteOrderAction(oRow, sLog, sOutputParams)
			Case "BROKERFILL"			
				iRetVal = FidessaBrokerFillAction(oRow, sLog, sOutputParams)
			Case "COMPLETEORDER"			
				iRetVal = FidessaCompleteOrderAction(oRow, sLog, sOutputParams)
			Case "VALIDATEORDERBOOKED"			
				iRetVal = FidessaValidateOrderBookedAction(oRow, sLog, sOutputParams)
			Case "HIGHLIGHTFIELDS"			
				iRetVal = FidessaHighlightFieldsAction(oRow, sLog, sOutputParams)
			Case "LOGOUT"			
				iRetVal = FidessaLogoutAction(oRow, sLog, sOutputParams)
			Case "STOP"	
				iRetVal = FidessaStopAction(oRow, sLog, sOutputParams)
			Case Else						
				iRetVal = XL_DISPATCH_UNKNOWN 	' Unknown keyword
		End Select

		RowDispatch = iRetVal
		oVBSFramework.oTraceLog.Exited(sRoutine & ", Keyword=""" & CStr(oRow.Cells(1, XL_KEYWORD).Value) & """")

	End Function
	
	'==============================================================================================
	' Function/Sub:
	'==============================================================================================
	Private Function FidessaStartAction(oRow, sLog, sOutputParams)
		Dim iRetVal, sApplication
		
		Const sRoutine = "clsFidessa.FidessaStartAction"
		oVBSFramework.oTraceLog.Entered(sRoutine)
		iRetVal = XL_DISPATCH_PASS
		
		'parameter handling
		sApplication = CStr(oRow.Cells(1, XL_PARM_001).Value)
		sDemoMode = UCase(CStr(oRow.Cells(1, XL_PARM_002).Value))
		LoadObjectRepository(sVBSFrameworkDir & "\ObjectMappings\ObjectRepositories\SimulatedApps.tsr")
		
		sLog = sLog & "I: Starting application: " & sApplication & vbLf

		SystemUtil.Run sApplication
		
		oVBSFramework.oTraceLog.Exited(sRoutine)
		FidessaStartAction = ListCurrentRow(oRow, sLog, sOutputParams)
	End Function

	'==============================================================================================
	' Function/Sub:
	'==============================================================================================
	Private Function FidessaHighlightFieldsAction(oRow, sLog, sOutputParams)
		Dim iRetVal
		
		Const sRoutine = "clsFidessa.FidessaHighlightFieldsAction"
		oVBSFramework.oTraceLog.Entered(sRoutine)
		iRetVal = XL_DISPATCH_PASS
		
		'parameter handling
		sLog = sLog & "I: Highlight fields" & vbLf
		
		Window("Simulated FIDESSA Application").Page("Simulated FIDESSA Application").WebEdit("UserId").Highlight
		Window("Simulated FIDESSA Application").Page("Simulated FIDESSA Application").WebEdit("Password").Highlight
		Window("Simulated FIDESSA Application").Page("Simulated FIDESSA Application").WebButton("Login").HighLight
		Window("Simulated FIDESSA Application").Page("Simulated FIDESSA Application").WebEdit("BuySell").Highlight
		Window("Simulated FIDESSA Application").Page("Simulated FIDESSA Application").WebEdit("Client").Highlight
		Window("Simulated FIDESSA Application").Page("Simulated FIDESSA Application").WebButton("Create").Highlight
		Window("Simulated FIDESSA Application").Page("Simulated FIDESSA Application").WebEdit("Action").Highlight
		Window("Simulated FIDESSA Application").Page("Simulated FIDESSA Application").WebButton("Action").Highlight		
		Window("Simulated FIDESSA Application").Page("Simulated FIDESSA Application").WebEdit("Output").Highlight

		oVBSFramework.oTraceLog.Exited(sRoutine)
		FidessaHighlightFieldsAction = ListCurrentRow(oRow, sLog, sOutputParams)
	End Function

	
	'==============================================================================================
	' Function/Sub:
	'==============================================================================================
	Private Function FidessaLoginAction(oRow, sLog, sOutputParams)
		Dim iRetVal, sLogin 
		
		Const sRoutine = "clsFidessa.FidessaLoginAction"
		oVBSFramework.oTraceLog.Entered(sRoutine)
		iRetVal = XL_DISPATCH_PASS
		
		'parameter handling
		sLogin = CStr(oRow.Cells(1, XL_PARM_001).Value)

		sLog = sLog & "I: Login in using: " & sLogin & vbLf
		
		Dim sUserId, sPassword
		sUserId = trim(split(split(sLogin,",")(1),"=")(1))
		sPassword = trim(split(split(sLogin,",")(2),"=")(1))
		If sDemoMode = "ON" Then Window("Simulated FIDESSA Application").Page("Simulated FIDESSA Application").WebEdit("UserId").Highlight
		Window("Simulated FIDESSA Application").Page("Simulated FIDESSA Application").WebEdit("UserId").Set sUserId
		If sDemoMode = "ON" Then Window("Simulated FIDESSA Application").Page("Simulated FIDESSA Application").WebEdit("Password").Highlight
		Window("Simulated FIDESSA Application").Page("Simulated FIDESSA Application").WebEdit("Password").Set sPassword
		If sDemoMode = "ON" Then Window("Simulated FIDESSA Application").Page("Simulated FIDESSA Application").WebButton("Login").HighLight
		Window("Simulated FIDESSA Application").Page("Simulated FIDESSA Application").WebButton("Login").Click
		
		If sDemoMode = "ON" Then Window("Simulated FIDESSA Application").Page("Simulated FIDESSA Application").WebEdit("Output").Highlight
		
		oVBSFramework.oTraceLog.Exited(sRoutine)
		FidessaLoginAction = ListCurrentRow(oRow, sLog, sOutputParams)
	End Function

	'==============================================================================================
	' Function/Sub:
	'==============================================================================================
	Private Function FidessaNewOrderAction(oRow, sLog, sOutputParams)
		Dim iRetVal, sBuySell, sOrderData
		
		Const sRoutine = "clsFidessa.FidessaNewOrderAction"
		oVBSFramework.oTraceLog.Entered(sRoutine)
		iRetVal = XL_DISPATCH_PASS
		
		'parameter handling
		sBuySell = CStr(oRow.Cells(1, XL_PARM_001).Value)
		sOrderData = CStr(oRow.Cells(1, XL_PARM_002).Value)
		If sDemoMode = "ON" Then Window("Simulated FIDESSA Application").Page("Simulated FIDESSA Application").WebEdit("BuySell").Highlight
		Window("Simulated FIDESSA Application").Page("Simulated FIDESSA Application").WebEdit("BuySell").Set sBuySell
		If sDemoMode = "ON" Then Window("Simulated FIDESSA Application").Page("Simulated FIDESSA Application").WebEdit("Client").Highlight
		Window("Simulated FIDESSA Application").Page("Simulated FIDESSA Application").WebEdit("Client").Set sOrderData
		If sDemoMode = "ON" Then Window("Simulated FIDESSA Application").Page("Simulated FIDESSA Application").WebButton("Create").Highlight
		Window("Simulated FIDESSA Application").Page("Simulated FIDESSA Application").WebButton("Create").Click

		If sDemoMode = "ON" Then Window("Simulated FIDESSA Application").Page("Simulated FIDESSA Application").WebEdit("Output").Highlight
		
		sLog = sLog & "I: " & sBuySell & "," & sOrderData & vbLf
		
		oVBSFramework.oTraceLog.Exited(sRoutine)
		FidessaNewOrderAction = ListCurrentRow(oRow, sLog, sOutputParams)
	End Function

	'==============================================================================================
	' Function/Sub:
	'==============================================================================================
	Private Function FidessaRouteOrderAction(oRow, sLog, sOutputParams)
		Dim iRetVal, sOrderData, sRouteTo
		
		Const sRoutine = "clsFidessa.FidessaRouteOrderAction"
		oVBSFramework.oTraceLog.Entered(sRoutine)
		iRetVal = XL_DISPATCH_PASS
		
		'parameter handling
		sOrderData = CStr(oRow.Cells(1, XL_PARM_001).Value)
		sRouteTo = CStr(oRow.Cells(1, XL_PARM_002).Value)
		
		
		If sDemoMode = "ON" Then Window("Simulated FIDESSA Application").Page("Simulated FIDESSA Application").WebEdit("Action").Highlight
		Window("Simulated FIDESSA Application").Page("Simulated FIDESSA Application").WebEdit("Action").Set CStr(oRow.Cells(1, XL_KEYWORD).Value) & ": " & sOrderData & "/" & sRouteTo
		If sDemoMode = "ON" Then Window("Simulated FIDESSA Application").Page("Simulated FIDESSA Application").WebButton("Action").Highlight		
		Window("Simulated FIDESSA Application").Page("Simulated FIDESSA Application").WebButton("Action").Click

		If sDemoMode = "ON" Then Window("Simulated FIDESSA Application").Page("Simulated FIDESSA Application").WebEdit("Output").Highlight
		
		sLog = sLog & "I: " & sOrderData & "," & sRouteTo & vbLf
		
		oVBSFramework.oTraceLog.Exited(sRoutine)
		FidessaRouteOrderAction = ListCurrentRow(oRow, sLog, sOutputParams)
	End Function
	
	
	'==============================================================================================
	' Function/Sub:
	'==============================================================================================
	Private Function FidessaBrokerFillAction(oRow, sLog, sOutputParams)
		Dim iRetVal, sAction
		
		Const sRoutine = "clsFidessa.FidessaBrokerFillAction"
		oVBSFramework.oTraceLog.Entered(sRoutine)
		iRetVal = XL_DISPATCH_PASS
		
		'parameter handling
		sAction = CStr(oRow.Cells(1, XL_PARM_001).Value)
		If sDemoMode = "ON" Then Window("Simulated FIDESSA Application").Page("Simulated FIDESSA Application").WebEdit("Action").Highlight
		Window("Simulated FIDESSA Application").Page("Simulated FIDESSA Application").WebEdit("Action").Set CStr(oRow.Cells(1, XL_KEYWORD).Value) & ": " & sAction
		If sDemoMode = "ON" Then Window("Simulated FIDESSA Application").Page("Simulated FIDESSA Application").WebButton("Action").Highlight		
		Window("Simulated FIDESSA Application").Page("Simulated FIDESSA Application").WebButton("Action").Click

		If sDemoMode = "ON" Then Window("Simulated FIDESSA Application").Page("Simulated FIDESSA Application").WebEdit("Output").Highlight
		
		sLog = sLog & "I: " & sAction & vbLf
		
		oVBSFramework.oTraceLog.Exited(sRoutine)
		FidessaBrokerFillAction = ListCurrentRow(oRow, sLog, sOutputParams)
	End Function
	
	
	'==============================================================================================
	' Function/Sub:
	'==============================================================================================
	Private Function FidessaCompleteOrderAction(oRow, sLog, sOutputParams)
		Dim iRetVal, sAction
		
		Const sRoutine = "clsFidessa.FidessaCompleteOrderAction"
		oVBSFramework.oTraceLog.Entered(sRoutine)
		iRetVal = XL_DISPATCH_PASS
		
		'parameter handling
		sAction = CStr(oRow.Cells(1, XL_PARM_001).Value)
		If sDemoMode = "ON" Then Window("Simulated FIDESSA Application").Page("Simulated FIDESSA Application").WebEdit("Action").Highlight
		Window("Simulated FIDESSA Application").Page("Simulated FIDESSA Application").WebEdit("Action").Set CStr(oRow.Cells(1, XL_KEYWORD).Value) & ": " & sAction
		If sDemoMode = "ON" Then Window("Simulated FIDESSA Application").Page("Simulated FIDESSA Application").WebButton("Action").Highlight		
		Window("Simulated FIDESSA Application").Page("Simulated FIDESSA Application").WebButton("Action").Click

		If sDemoMode = "ON" Then Window("Simulated FIDESSA Application").Page("Simulated FIDESSA Application").WebEdit("Output").Highlight
		
		sLog = sLog & "I: " & sAction & vbLf
		
		oVBSFramework.oTraceLog.Exited(sRoutine)
		FidessaCompleteOrderAction = ListCurrentRow(oRow, sLog, sOutputParams)
	End Function
	
	
		'==============================================================================================
	' Function/Sub:
	'==============================================================================================
	Private Function FidessaValidateOrderBookedAction(oRow, sLog, sOutputParams)
		Dim iRetVal, sAction
		
		Const sRoutine = "clsFidessa.FidessaValidateOrderBookedAction"
		oVBSFramework.oTraceLog.Entered(sRoutine)
		iRetVal = XL_DISPATCH_PASS
		
		'parameter handling
		sAction = CStr(oRow.Cells(1, XL_PARM_001).Value)
		If sDemoMode = "ON" Then Window("Simulated FIDESSA Application").Page("Simulated FIDESSA Application").WebEdit("Action").Highlight
		Window("Simulated FIDESSA Application").Page("Simulated FIDESSA Application").WebEdit("Action").Set CStr(oRow.Cells(1, XL_KEYWORD).Value) & ": " & sAction
		If sDemoMode = "ON" Then Window("Simulated FIDESSA Application").Page("Simulated FIDESSA Application").WebButton("Action").Highlight		
		Window("Simulated FIDESSA Application").Page("Simulated FIDESSA Application").WebButton("Action").Click

		If sDemoMode = "ON" Then Window("Simulated FIDESSA Application").Page("Simulated FIDESSA Application").WebEdit("Output").Highlight
		
		sLog = sLog & "I: " & sAction & vbLf
		
		oVBSFramework.oTraceLog.Exited(sRoutine)
		FidessaValidateOrderBookedAction = ListCurrentRow(oRow, sLog, sOutputParams)
	End Function

	'==============================================================================================
	' Function/Sub:
	'==============================================================================================
	Private Function FidessaLogoutAction(oRow, sLog, sOutputParams)
		Dim iRetVal 
		
		Const sRoutine = "clsFidessa.FidessaLogoutAction"
		oVBSFramework.oTraceLog.Entered(sRoutine)
		iRetVal = XL_DISPATCH_PASS
		
		sLog = sLog & "I: Logout ..." & vbLf
		'call PopUp("Logged out", 1, "Logout")

		oVBSFramework.oTraceLog.Exited(sRoutine)
		FidessaLogoutAction = ListCurrentRow(oRow, sLog, sOutputParams)
	End Function
	
	'==============================================================================================
	' Function/Sub:
	'==============================================================================================
	Private Function FidessaStopAction(oRow, sLog, sOutputParams)
		Dim iRetVal 
		
		Const sRoutine = "clsFidessa.FidessaStopAction"
		oVBSFramework.oTraceLog.Entered(sRoutine)
		iRetVal = XL_DISPATCH_PASS
		
		If sDemoMode = "ON" Then Window("Simulated FIDESSA Application").Page("Simulated FIDESSA Application").WebButton("Exit").Highlight
		Window("Simulated FIDESSA Application").Page("Simulated FIDESSA Application").WebButton("Exit").Click
		sLog = sLog & "I: closing Fidessa" & vbLf
	
		oVBSFramework.oTraceLog.Exited(sRoutine)
		FidessaStopAction = ListCurrentRow(oRow, sLog, sOutputParams)
	End Function
	
	'==============================================================================================
	' Function/Sub:
	'==============================================================================================
	Private Function ListCurrentRow(oRow, sLog, sOutputParams)
		Dim iRetVal, sParam
		
		Const sRoutine = "clsFidessa.ListCurrentRow"
		oVBSFramework.oTraceLog.Entered(sRoutine)
		iRetVal = XL_DISPATCH_PASS
		
		If sDemoMode = "ON" Then 
			Dim i, sRowInput
			sRowInput = CStr(oRow.Cells(1, XL_KEYWORD).Value) & vbLF
			for i = XL_PARM_001 to XL_PARM_010
				sParam = CStr(oRow.Cells(1, i).Value)
				if trim(sParam) = "" then exit for
				sRowInput = sRowInput & "Input " & i - XL_PARM_001 + 1 & ":" & sParam & vbNewLine
			next
			
			PopUp sRowInput, 1, "Row:" & oRow.Cells(1, i).row
			'msgbox sRowInput
		End if
		oVBSFramework.oTraceLog.Exited(sRoutine)
		ListCurrentRow = iRetVal
	End Function

'--------------------------------------------------------------------------------------------------
' Class End clsFidessa
'--------------------------------------------------------------------------------------------------
end class

Public oFidessa
Set oFidessa = new clsFidessa
oVBSFramework.oTestObjects.Add "FIDESSA", oFidessa

