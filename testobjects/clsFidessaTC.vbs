Option Explicit
'==================================================================================================
' clsFidessaTC.vbs
'
' Purpose
'==================================================================================================
'--------------------------------------------------------------------------------------------------
' Constants
'--------------------------------------------------------------------------------------------------

'--------------------------------------------------------------------------------------------------
' Class Start
'--------------------------------------------------------------------------------------------------

Class clsFidessaTC

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

		Const sRoutine = "clsFidessaTC.RowDispatch"
		oVBSFramework.oTraceLog.Entered(sRoutine & ", Keyword=""" & CStr(oRow.Cells(1, XL_KEYWORD).Value) & """")

		iRetVal = XL_DISPATCH_PASS
		
		' Get keyword from spreadsheet as string
		sKeyword = CStr(oRow.Cells(1, XL_KEYWORD).Value)
		sKeyword = right(sKeyword, len(sKeyword)-inStrRev(sKeyword, "."))

		Select Case uCase(sKeyword)
			Case "START"			
				iRetVal = FidessaTCStartAction(oRow, sLog, sOutputParams)
			Case "LOGIN"			
				iRetVal = FidessaTCLoginAction(oRow, sLog, sOutputParams)
			Case "NEWORDER"			
				iRetVal = FidessaTCNewOrderAction(oRow, sLog, sOutputParams)
			Case "ROUTEORDER"			
				iRetVal = FidessaTCRouteOrderAction(oRow, sLog, sOutputParams)
			Case "BROKERFILL"			
				iRetVal = FidessaTCBrokerFillAction(oRow, sLog, sOutputParams)
			Case "COMPLETEORDER"			
				iRetVal = FidessaTCCompleteOrderAction(oRow, sLog, sOutputParams)
			Case "VALIDATEORDERBOOKED"			
				iRetVal = FidessaTCValidateOrderBookedAction(oRow, sLog, sOutputParams)
			Case "HIGHLIGHTFIELDS"			
				iRetVal = FidessaTCHighlightFieldsAction(oRow, sLog, sOutputParams)
			Case "LOGOUT"			
				iRetVal = FidessaTCLogoutAction(oRow, sLog, sOutputParams)
			Case "STOP"	
				iRetVal = FidessaTCStopAction(oRow, sLog, sOutputParams)
			Case Else						
				iRetVal = XL_DISPATCH_UNKNOWN 	' Unknown keyword
		End Select

		RowDispatch = iRetVal
		oVBSFramework.oTraceLog.Exited(sRoutine & ", Keyword=""" & CStr(oRow.Cells(1, XL_KEYWORD).Value) & """")

	End Function
	
	'==============================================================================================
	' Function/Sub:
	'==============================================================================================
	Private Function FidessaTCStartAction(oRow, sLog, sOutputParams)
		Dim iRetVal, sApplication
		
		Const sRoutine = "clsFidessaTC.FidessaTCStartAction"
		oVBSFramework.oTraceLog.Entered(sRoutine)
		iRetVal = XL_DISPATCH_PASS
		
		'parameter handling
		sApplication = CStr(oRow.Cells(1, XL_PARM_001).Value)
		sDemoMode = UCase(CStr(oRow.Cells(1, XL_PARM_002).Value))
		
		sLog = sLog & "I: Starting application: " & sApplication & vbLf

		'SystemUtil.Run sApplication
		Call CreateObject("Wscript.Shell").Run(sApplication)
		
		oVBSFramework.oTraceLog.Exited(sRoutine)
		FidessaTCStartAction = ListCurrentRow(oRow, sLog, sOutputParams)
	End Function

	'==============================================================================================
	' Function/Sub:
	'==============================================================================================
	Private Function FidessaTCHighlightFieldsAction(oRow, sLog, sOutputParams)
		Dim iRetVal
		
		Const sRoutine = "clsFidessaTC.FidessaTCHighlightFieldsAction"
		oVBSFramework.oTraceLog.Entered(sRoutine)
		iRetVal = XL_DISPATCH_PASS
		
		'parameter handling
		sLog = sLog & "I: Highlight fields" & vbLf
		
		'Window("Simulated FidessaTC Application").Page("Simulated FidessaTC Application").WebEdit("UserId").Highlight
		'Window("Simulated FidessaTC Application").Page("Simulated FidessaTC Application").WebEdit("Password").Highlight
		'Window("Simulated FidessaTC Application").Page("Simulated FidessaTC Application").WebButton("Login").HighLight
		'Window("Simulated FidessaTC Application").Page("Simulated FidessaTC Application").WebEdit("BuySell").Highlight
		'Window("Simulated FidessaTC Application").Page("Simulated FidessaTC Application").WebEdit("Client").Highlight
		'Window("Simulated FidessaTC Application").Page("Simulated FidessaTC Application").WebButton("Create").Highlight
		'Window("Simulated FidessaTC Application").Page("Simulated FidessaTC Application").WebEdit("Action").Highlight
		'Window("Simulated FidessaTC Application").Page("Simulated FidessaTC Application").WebButton("Action").Highlight		
		'Window("Simulated FidessaTC Application").Page("Simulated FidessaTC Application").WebEdit("Output").Highlight

		oVBSFramework.oTraceLog.Exited(sRoutine)
		FidessaTCHighlightFieldsAction = ListCurrentRow(oRow, sLog, sOutputParams)
	End Function

	
	'==============================================================================================
	' Function/Sub:
	'==============================================================================================
	Private Function FidessaTCLoginAction(oRow, sLog, sOutputParams)
		Dim iRetVal, sLogin 
		
		Const sRoutine = "clsFidessaTC.FidessaTCLoginAction"
		oVBSFramework.oTraceLog.Entered(sRoutine)
		iRetVal = XL_DISPATCH_PASS
		
		'parameter handling
		sLogin = CStr(oRow.Cells(1, XL_PARM_001).Value)

		sLog = sLog & "I: Login in using: " & sLogin & vbLf
		
		Dim sUserId, sPassword
		sUserId = trim(split(split(sLogin,",")(1),"=")(1))
		sPassword = trim(split(split(sLogin,",")(2),"=")(1))
		
		Dim page
		Set page = Aliases.mshta.wndHTMLApplicationHostWindowClass.browser.pageSimulatedFidessaApplication
		
		If sDemoMode = "ON" Then Sys.HighlightObject page.textboxTxtuserid, 5, 255
		Call page.textboxTxtuserid.SetText(sUserId)
		If sDemoMode = "ON" Then Sys.HighlightObject page.passwordboxTxtpassword, 5, 255
		Call page.passwordboxTxtpassword.SetText(sPassword)
		If sDemoMode = "ON" Then Sys.HighlightObject page.buttonLogin, 5, 255
		page.buttonLogin.ClickButton
		
		If sDemoMode = "ON" Then Sys.HighlightObject page.textareaTxtoutput, 5, 255
		
		oVBSFramework.oTraceLog.Exited(sRoutine)
		FidessaTCLoginAction = ListCurrentRow(oRow, sLog, sOutputParams)
	End Function

	'==============================================================================================
	' Function/Sub:
	'==============================================================================================
	Private Function FidessaTCNewOrderAction(oRow, sLog, sOutputParams)
		Dim iRetVal, sBuySell, sOrderData
		
		Const sRoutine = "clsFidessaTC.FidessaTCNewOrderAction"
		oVBSFramework.oTraceLog.Entered(sRoutine)
		iRetVal = XL_DISPATCH_PASS
		
		'parameter handling
		sBuySell = CStr(oRow.Cells(1, XL_PARM_001).Value)
		sOrderData = CStr(oRow.Cells(1, XL_PARM_002).Value)
		
		
		Dim page
		Set page = Aliases.mshta.wndHTMLApplicationHostWindowClass.browser.pageSimulatedFidessaApplication
		
		If sDemoMode = "ON" Then Sys.HighlightObject page.textboxTxtbuysell, 5, 255
		Call page.textboxTxtbuysell.SetText(sBuySell)
		If sDemoMode = "ON" Then Sys.HighlightObject page.textboxTxtclient, 5, 255
		Call page.textboxTxtclient.SetText(sOrderData)
		If sDemoMode = "ON" Then Sys.HighlightObject page.buttonCreate, 5, 255
		page.buttonCreate.ClickButton

		If sDemoMode = "ON" Then Sys.HighlightObject page.textareaTxtoutput, 5, 255
		
		sLog = sLog & "I: " & sBuySell & "," & sOrderData & vbLf
		
		oVBSFramework.oTraceLog.Exited(sRoutine)
		FidessaTCNewOrderAction = ListCurrentRow(oRow, sLog, sOutputParams)
	End Function

	'==============================================================================================
	' Function/Sub:
	'==============================================================================================
	Private Function FidessaTCRouteOrderAction(oRow, sLog, sOutputParams)
		Dim iRetVal, sOrderData, sRouteTo
		
		Const sRoutine = "clsFidessaTC.FidessaTCRouteOrderAction"
		oVBSFramework.oTraceLog.Entered(sRoutine)
		iRetVal = XL_DISPATCH_PASS
		
		'parameter handling
		sOrderData = CStr(oRow.Cells(1, XL_PARM_001).Value)
		sRouteTo = CStr(oRow.Cells(1, XL_PARM_002).Value)
		
		Dim page
		Set page = Aliases.mshta.wndHTMLApplicationHostWindowClass.browser.pageSimulatedFidessaApplication
		
		If sDemoMode = "ON" Then Sys.HighlightObject page.textboxTxtaction, 5, 255
		Call page.textboxTxtaction.SetText(CStr(oRow.Cells(1, XL_KEYWORD).Value) & ": " & sOrderData & "/" & sRouteTo)
		If sDemoMode = "ON" Then Sys.HighlightObject page.buttonAction, 5, 255
		page.buttonAction.ClickButton

		If sDemoMode = "ON" Then Sys.HighlightObject page.textareaTxtoutput, 5, 255
		
		sLog = sLog & "I: " & sOrderData & "," & sRouteTo & vbLf
		
		oVBSFramework.oTraceLog.Exited(sRoutine)
		FidessaTCRouteOrderAction = ListCurrentRow(oRow, sLog, sOutputParams)
	End Function
	
	
	'==============================================================================================
	' Function/Sub:
	'==============================================================================================
	Private Function FidessaTCBrokerFillAction(oRow, sLog, sOutputParams)
		Dim iRetVal, sAction
		
		Const sRoutine = "clsFidessaTC.FidessaTCBrokerFillAction"
		oVBSFramework.oTraceLog.Entered(sRoutine)
		iRetVal = XL_DISPATCH_PASS
		
		'parameter handling
		sAction = CStr(oRow.Cells(1, XL_PARM_001).Value)

		Dim page
		Set page = Aliases.mshta.wndHTMLApplicationHostWindowClass.browser.pageSimulatedFidessaApplication
		
		If sDemoMode = "ON" Then Sys.HighlightObject page.textboxTxtaction, 5, 255
		Call page.textboxTxtaction.SetText(CStr(oRow.Cells(1, XL_KEYWORD).Value) & ": " & sAction)
		If sDemoMode = "ON" Then Sys.HighlightObject page.buttonAction, 5, 255
		page.buttonAction.ClickButton

		If sDemoMode = "ON" Then Sys.HighlightObject page.textareaTxtoutput, 5, 255
		
		sLog = sLog & "I: " & sAction & vbLf
		
		oVBSFramework.oTraceLog.Exited(sRoutine)
		FidessaTCBrokerFillAction = ListCurrentRow(oRow, sLog, sOutputParams)
	End Function
	
	
	'==============================================================================================
	' Function/Sub:
	'==============================================================================================
	Private Function FidessaTCCompleteOrderAction(oRow, sLog, sOutputParams)
		Dim iRetVal, sAction
		
		Const sRoutine = "clsFidessaTC.FidessaTCCompleteOrderAction"
		oVBSFramework.oTraceLog.Entered(sRoutine)
		iRetVal = XL_DISPATCH_PASS
		
		'parameter handling
		sAction = CStr(oRow.Cells(1, XL_PARM_001).Value)
		Dim page
		Set page = Aliases.mshta.wndHTMLApplicationHostWindowClass.browser.pageSimulatedFidessaApplication
		
		If sDemoMode = "ON" Then Sys.HighlightObject page.textboxTxtaction, 5, 255
		Call page.textboxTxtaction.SetText(CStr(oRow.Cells(1, XL_KEYWORD).Value) & ": " & sAction)
		If sDemoMode = "ON" Then Sys.HighlightObject page.buttonAction, 5, 255
		page.buttonAction.ClickButton

		If sDemoMode = "ON" Then Sys.HighlightObject page.textareaTxtoutput, 5, 255
		
		sLog = sLog & "I: " & sAction & vbLf
		
		oVBSFramework.oTraceLog.Exited(sRoutine)
		FidessaTCCompleteOrderAction = ListCurrentRow(oRow, sLog, sOutputParams)
	End Function
	
	
		'==============================================================================================
	' Function/Sub:
	'==============================================================================================
	Private Function FidessaTCValidateOrderBookedAction(oRow, sLog, sOutputParams)
		Dim iRetVal, sAction
		
		Const sRoutine = "clsFidessaTC.FidessaTCValidateOrderBookedAction"
		oVBSFramework.oTraceLog.Entered(sRoutine)
		iRetVal = XL_DISPATCH_PASS
		
		'parameter handling
		sAction = CStr(oRow.Cells(1, XL_PARM_001).Value)
		Dim page
		Set page = Aliases.mshta.wndHTMLApplicationHostWindowClass.browser.pageSimulatedFidessaApplication
		
		If sDemoMode = "ON" Then Sys.HighlightObject page.textboxTxtaction, 5, 255
		Call page.textboxTxtaction.SetText(CStr(oRow.Cells(1, XL_KEYWORD).Value) & ": " & sAction)
		If sDemoMode = "ON" Then Sys.HighlightObject page.buttonAction, 5, 255
		page.buttonAction.ClickButton

		If sDemoMode = "ON" Then Sys.HighlightObject page.textareaTxtoutput, 5, 255
		
		sLog = sLog & "I: " & sAction & vbLf
		
		oVBSFramework.oTraceLog.Exited(sRoutine)
		FidessaTCValidateOrderBookedAction = ListCurrentRow(oRow, sLog, sOutputParams)
	End Function

	'==============================================================================================
	' Function/Sub:
	'==============================================================================================
	Private Function FidessaTCLogoutAction(oRow, sLog, sOutputParams)
		Dim iRetVal 
		
		Const sRoutine = "clsFidessaTC.FidessaTCLogoutAction"
		oVBSFramework.oTraceLog.Entered(sRoutine)
		iRetVal = XL_DISPATCH_PASS
		
		Dim page
		Set page = Aliases.mshta.wndHTMLApplicationHostWindowClass.browser.pageSimulatedFidessaApplication

		If sDemoMode = "ON" Then Sys.HighlightObject page.buttonLogout, 5, 255
		page.buttonLogout.ClickButton
		
		sLog = sLog & "I: Logout ..." & vbLf
		'call PopUp("Logged out", 1, "Logout")

		oVBSFramework.oTraceLog.Exited(sRoutine)
		FidessaTCLogoutAction = ListCurrentRow(oRow, sLog, sOutputParams)
	End Function
	
	'==============================================================================================
	' Function/Sub:
	'==============================================================================================
	Private Function FidessaTCStopAction(oRow, sLog, sOutputParams)
		Dim iRetVal 
		
		Const sRoutine = "clsFidessaTC.FidessaTCStopAction"
		oVBSFramework.oTraceLog.Entered(sRoutine)
		iRetVal = XL_DISPATCH_PASS
		
		Dim page
		Set page = Aliases.mshta.wndHTMLApplicationHostWindowClass.browser.pageSimulatedFidessaApplication

		If sDemoMode = "ON" Then Sys.HighlightObject page.buttonExit, 5, 255
		page.buttonExit.ClickButton

		sLog = sLog & "I: closing FidessaTC" & vbLf
	
		oVBSFramework.oTraceLog.Exited(sRoutine)
		FidessaTCStopAction = ListCurrentRow(oRow, sLog, sOutputParams)
	End Function
	
	'==============================================================================================
	' Function/Sub:
	'==============================================================================================
	Private Function ListCurrentRow(oRow, sLog, sOutputParams)
		Dim iRetVal, sParam
		
		Const sRoutine = "clsFidessaTC.ListCurrentRow"
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
' Class End clsFidessaTC
'--------------------------------------------------------------------------------------------------
end class

Public oFidessaTC
Set oFidessaTC = new clsFidessaTC
oVBSFramework.oTestObjects.Add "FidessaTC", oFidessaTC

