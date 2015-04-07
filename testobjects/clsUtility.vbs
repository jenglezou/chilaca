Option Explicit
'==================================================================================================
' clsUtility.vbs
'
'==================================================================================================
'
'--------------------------------------------------------------------------------------------------
' AMENDEMENT HISTORY
'--------------------------------------------------------------------------------------------------
' Reason:
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
' DESCRIPTION
'--------------------------------------------------------------------------------------------------
'**
' Describe change detail, include name of function amended
'**
'--------------------------------------------------------------------------------------------------
'
'on error resume next

'--------------------------------------------------------------------------------------------------
' Constants
'--------------------------------------------------------------------------------------------------

'--------------------------------------------------------------------------------------------------
' Class Start
'--------------------------------------------------------------------------------------------------

Class clsUtility
	'==============================================================================================
	' CLASS VARIABLES
	'==============================================================================================

	'----------------------------------------------------------------------------------------------
	' Public Variables
	'----------------------------------------------------------------------------------------------
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
	' Function:		RowDispatch
	' Purpose:		Determines which function to call based on keyword
	'
	' Parameters:	oRow - Range object containing the current row
	'
	' Returns:		Return code from keyword function or XL_DISPATCH_UNKNOWN if
	'				keyword not supported
	'
	'==============================================================================================
	Public Function RowDispatch(oRow, sLog, sOutputParams)
		Dim iRetVal, sKeyword
		'Dim sApp,iAppLen,oApp,sTopRow,HeaderArrayFlag

		Const sRoutine = "clsUtility.RowDispatch"
		oVBSFramework.oTraceLog.Entered(sRoutine & ", Keyword=" & CStr(oRow.Cells(1, XL_KEYWORD).Value))

		' Get keyword from spreadsheet as string
		sKeyword = CStr(oRow.Cells(1, XL_KEYWORD).Value)
		sKeyword = right(sKeyword, len(sKeyword)-inStrRev(sKeyword, "."))

		Select Case UCase(sKeyword)
			'Dispatch known keyword functions
			'Case "<keyword>"				iRetVal = K_<Keyword>( oRow )
			Case "SELECTITEM"				
				iRetVal = SelectItem(oRow)
			Case "SELECTOPTION"				
				iRetVal = SelectOption(oRow)
			Case "PAUSE", "PAUSESHEET", "SKIPSHEET"
				iRetVal = PauseSheet(oRow)
			Case "PAUSEROW", "SKIPROW"		
				iRetVal = PauseRow(oRow)
			Case "TEMP"						
				iRetVal = temp(oRow)
			Case "CONTINUEIFTRUE", "RUNIFTRUE"
				iRetVal = ContinueIfTrue(oRow)				
			Case Else						
			iRetVal = XL_DISPATCH_UNKNOWN 	' Unknown keyword
		End Select

		RowDispatch = iRetVal
		oVBSFramework.oTraceLog.Exited(sRoutine & ", Keyword=""" & CStr(oRow.Cells(1, XL_KEYWORD).Value) & """")
	End Function
	
	'==============================================================================================
	' Function:		ContinueIfTrue
	' Purpose:		
	'
	' Parameters:	
	'				
	'
	' Returns:		XL_DISPATCH_END, XL_DISPATCH_PASS
	'
	'==============================================================================================
	Public Function ContinueIfTrue(oRow)
		iRetVal = XL_DISPATCH_PASS
		if not UCase(CStr(oRow.Cells(1, XL_PARM_001).Value)) = "TRUE" then
			iRetVal = XL_DISPATCH_END
		end if
		ContinueIfTrue = iRetVal
	End Function
	'==============================================================================================
	' Function:		SelectItem
	' Purpose:		
	'
	' Parameters:	Parm001 - Is a prompt
	'				Parm002 - Is a comma seperated list values
	'
	' Returns:		XL_DISPATCH_FAIL, XL_DISPATCH_PASS
	'
	'==============================================================================================
	Public Function SelectItem(oRow)
		Dim iRetVal, sPrompt, sItemList, arrItems
		Dim sItem, sSelectedItem, vbButtonPressed
		Dim fso, sTmpFolder, oTmpFolder, oItem
		Dim sCurrentTestDir,sCodePath
		Dim oDialog, iYesNo
		Dim oItemList, bItemSelected

		Const sRoutine = "clsUtility.SelectItem"
		oVBSFramework.oTraceLog.Entered(sRoutine & ", Keyword=" & CStr(oRow.Cells(1, XL_KEYWORD).Value))
		
		iYesNo = vbYes
		vbButtonPressed = -1

		' Get parameters from spreadsheet as strings
		sPrompt = CStr(oRow.Cells(1, XL_PARM_001).Value)
		sItemList = CStr(oRow.Cells(1, XL_PARM_002).Value)
		arrItems = Split(sItemList,",",-1,1)
		sSelectedItem = ""

		on error resume next
		set oItemList = createobject("JEUtilities.itemlist")
		if oItemList is Nothing then
			On Error GoTo 0	'Turn off on error
			'msgbox "JEUtilies.exe is not loaded." & chr(13) & "Will use the open file dialog instead."
			'Use open dialog
			sTmpFolder = "c:\tmpselectitem"
			Set fso = CreateObject("Scripting.FileSystemObject")

			'Create folder
			on error resume next
			If (fso.FolderExists(sTmpFolder)) Then
				fso.DeleteFolder(sTmpFolder)
			End If

			fso.CreateFolder(sTmpFolder)

			For Each sItem in arrItems
			   fso.OpenTextFile sTmpFolder & "\" & trim(sItem), 2, True
			Next
			On Error GoTo 0	'Turn off on error


			bItemSelected = False
			Set oDialog = CreateObject("UserAccounts.CommonDialog")
			do while bItemSelected = False
				'oDialog.DialogTitle = "Select Item"
				oDialog.InitialDir = sTmpFolder
				'sQTPWinState = oQTPApp.WindowState
				'oQTPApp.WindowState = "Normal"
				bItemSelected = oDialog.ShowOpen
				if bItemSelected = False then
					iYesNo = msgbox("You did not select an Item." & Chr(13) & "Press Yes to try again or No to stop the test.", vbYesNo, "Try again.")
					if iYesNo = vbNo then
						exit do
					end if
				end if
			loop

			sSelectedItem = fso.GetFileName(oDialog.FileName)
			'oQTPApp.WindowState = sQTPWinState
			Set fso = Nothing
			set oDialog = Nothing

		else	'Use JEUtilities.exe
			oItemList.reset
			For Each sItem in arrItems
				oItemList.AddItem trim(sItem)
			Next

			oItemList.Caption = sPrompt
			oItemList.Description = ""
			
			oItemList.show
			sSelectedItem = oItemList.SelectedItem
			vbButtonPressed = oItemList.ButtonPressed
			Do while sSelectedItem = ""
				iYesNo = msgbox("You did not select or enter an item." & Chr(13) & "Press Yes to try again or No to stop the test.", vbYesNo, "Try again.")
				if iYesNo = vbNo then
					exit do
				end if
				oItemList.show
				sSelectedItem = oItemList.SelectedItem
				vbButtonPressed = oItemList.ButtonPressed
			loop

			Set oItemList = Nothing
		end if

		oRow.Cells(1, XL_OUTPUT).Value = UCase(sSelectedItem)

		iRetVal = XL_DISPATCH_PASS
		if sSelectedItem = "" then
			iRetVal = XL_DISPATCH_CANCEL
		end if

		' Set return value
		SelectItem = iRetVal
		oVBSFramework.oTraceLog.Exited(sRoutine & ", Keyword=""" & CStr(oRow.Cells(1, XL_KEYWORD).Value) & """")
	End Function

	'==============================================================================================
	' Function:		SelectOption
	' Purpose:		
	'
	' Parameters:	Parm001 - Is a prompt
	'				Parm002 - Is a comma seperated list values
	'
	' Returns:		XL_DISPATCH_FAIL, XL_DISPATCH_PASS
	'
	'==============================================================================================
	Public Function SelectOption(oRow)
		Dim iRetVal, sPrompt, sButtonList, arrButtons
		Dim sButton, sSelectedButton, vbButtonPressed
		Dim iYesNo
		Dim oButtonList, bButtonSelected

		Const sRoutine = "clsUtility.SelectOption"
		oVBSFramework.oTraceLog.Entered(sRoutine & ", Keyword=" & CStr(oRow.Cells(1, XL_KEYWORD).Value))

		iYesNo = vbYes
		vbButtonPressed = -1

		' Get parameters from spreadsheet as strings
		sPrompt = CStr(oRow.Cells(1, XL_PARM_001).Value)
		sButtonList = CStr(oRow.Cells(1, XL_PARM_002).Value)
		arrButtons = Split(sButtonList,",",-1,1)
		sSelectedButton = ""

		on error resume next
		set oButtonList = createobject("JEUtilities.ButtonList")
		if oButtonList is Nothing then
			On Error GoTo 0	'Turn off on error
			'msgbox "JEUtilies.exe is not loaded." & chr(13) & "Will use fallback method instead."
			'Use fallback - only the first two options are used
			Select Case Msgbox("Press Yes for " & arrButtons(0) & _
								", No for " & arrButtons(1) & " or Cancel to end the Test.", _
								vbYesNoCancel, sPrompt)
				Case vbYes
					oRow.Cells(1, XL_OUTPUT).Value = ucase(trim(arrButtons(0)))
					sSelectedButton = ucase(trim(arrButtons(0)))
					iRetVal = XL_DISPATCH_PASS
				Case vbNo
					oRow.Cells(1, XL_OUTPUT).Value = ucase(trim(arrButtons(1)))
					sSelectedButton = ucase(trim(arrButtons(1)))
					iRetVal = XL_DISPATCH_PASS
				Case vbCancel	iRetVal = XL_DISPATCH_CANCEL
			End Select

		else	'Use JEUtilities.exe
			oButtonList.reset
			For Each sButton in arrButtons
				oButtonList.AddButton trim(sButton)
			Next

			oButtonList.Caption = sPrompt
			oButtonList.Description = ""
			oButtonList.show
			sSelectedButton = oButtonList.SelectedButton
			vbButtonPressed = oButtonList.ButtonPressed
			Do while sSelectedButton = ""
				iYesNo = msgbox("You did not select an option." & Chr(13) & "Press Yes to try again or No to stop the test.", vbYesNo, "Try again.")
				if iYesNo = vbNo then
					exit do
				end if
				oButtonList.show
				sSelectedButton = oButtonList.SelectedButton
				vbButtonPressed = oButtonList.ButtonPressed
			loop

			Set oButtonList = Nothing
		end if

		oRow.Cells(1, XL_OUTPUT).Value = UCase(sSelectedButton)

		iRetVal = XL_DISPATCH_PASS
		if sSelectedButton = "" then
			iRetVal = XL_DISPATCH_CANCEL
		end if

		' Set return value
		SelectOption = iRetVal
		oVBSFramework.oTraceLog.Exited(sRoutine & ", Keyword=""" & CStr(oRow.Cells(1, XL_KEYWORD).Value) & """")
	End Function

	'==============================================================================================
	' Function:		PauseSheet
	' Purpose:		
	'
	' Parameters:	Parm001 - Is a prompt
	'				Parm002 - Is a comma seperated list values
	'
	' Returns:		XL_DISPATCH_FAIL, XL_DISPATCH_PASS
	'
	'==============================================================================================
	Public Function PauseSheet(oRow)
		Dim iRetVal
		Dim sSelectedButton, vbButtonPressed
		Dim iYesNo
		Dim oButtonList
		
		Const sRUN = "Run this sheet"
		Const sSKIP = "Skip this sheet"
		Const sEND = "End the test"

		Const sRoutine = "clsUtility.PauseSheet"
		oVBSFramework.oTraceLog.Entered(sRoutine & ", Keyword=" & CStr(oRow.Cells(1, XL_KEYWORD).Value))

		'iRetVal = XL_DISPATCH_SKIP		'By default run the sheet
		iRetVal = XL_DISPATCH_CANCEL		'By default end the test
		iYesNo = vbYes
		vbButtonPressed = -1
		sSelectedButton = ""

		on error resume next
		set oButtonList = createobject("JEUtilities.ButtonList")
		if oButtonList is Nothing then
			On Error GoTo 0	'Turn off on error
			'msgbox "JEUtilies.exe is not loaded." & chr(13) & "Will use fallback method instead."
			'Use fallback - only the first two options are used

			Select Case Msgbox("Press Yes to continue with this sheet, No to end this Sheet or Cancel to end the Test.", _
								vbYesNoCancel, "Continue Sheet - " & oRow.Parent.Name)
				Case vbYes		iRetVal = XL_DISPATCH_SKIP
				Case vbNo		iRetVal = XL_DISPATCH_END
				Case vbCancel	iRetVal = XL_DISPATCH_CANCEL
			End Select

		else	'Use JEUtilities.exe
			oButtonList.reset
			oButtonList.AddButton sRUN
			oButtonList.AddButton sSKIP
			oButtonList.AddButton sEND
			oButtonList.Caption = "Paused at sheet"
			oButtonList.description = oRow.Parent.Name
			oButtonList.show
			sSelectedButton = oButtonList.SelectedButton
			vbButtonPressed = oButtonList.ButtonPressed
			Do while sSelectedButton = ""
				iYesNo = msgbox("You did not select an option." & Chr(13) & "Press Yes to try again or No to stop the test.", vbYesNo, "Try again.")
				if iYesNo = vbNo then
					exit do
				end if
				oButtonList.show
				sSelectedButton = oButtonList.SelectedButton
				vbButtonPressed = oButtonList.ButtonPressed
			loop

			Select Case sSelectedButton
				Case sRUN		iRetVal = XL_DISPATCH_SKIP
				Case sSKIP		iRetVal = XL_DISPATCH_END
				Case sEND		iRetVal = XL_DISPATCH_CANCEL
			End Select

			Set oButtonList = Nothing
		end if

		' Set return value
		PauseSheet = iRetVal
		oVBSFramework.oTraceLog.Exited(sRoutine & ", Keyword=""" & CStr(oRow.Cells(1, XL_KEYWORD).Value) & """")
	End Function
	
	'==============================================================================================
	' Function:		PauseRow
	' Purpose:		
	'
	' Parameters:	Parm001 - Is a prompt
	'				Parm002 - Is a comma seperated list values
	'
	' Returns:		XL_DISPATCH_FAIL, XL_DISPATCH_PASS
	'
	'==============================================================================================
	Public Function PauseRow(oRow)
		Dim iRetVal
		Dim sSelectedButton, vbButtonPressed
		Dim iYesNo
		Dim oButtonList
		
		Const sRUN = "Run the next row"
		Const sSKIP = "Skip the next row"
		Const sEND = "End the test"

		Const sRoutine = "clsUtility.PauseRow"
		oVBSFramework.oTraceLog.Entered(sRoutine & ", Keyword=" & CStr(oRow.Cells(1, XL_KEYWORD).Value))

		'iRetVal = XL_DISPATCH_SKIP		'By default run the row
		iRetVal = XL_DISPATCH_CANCEL		'By default end the test
		iYesNo = vbYes
		vbButtonPressed = -1
		sSelectedButton = ""

		on error resume next
		set oButtonList = createobject("JEUtilities.ButtonList")
		if oButtonList is Nothing then
			On Error GoTo 0	'Turn off on error
			'msgbox "JEUtilies.exe is not loaded." & chr(13) & "Will use fallback method instead."
			'Use fallback - only the first two options are used
			Select Case Msgbox("Press Yes to run the next row, No to skip the next step or Cancel to end the Test.", _
								vbYesNoCancel, "Run next row - " & oRow.Cells(2, XL_KEYWORD).Value)
			Case vbYes		iRetVal = XL_DISPATCH_SKIP
			Case vbNo		oRow.Cells(2, XL_DISABLE).Value = "TRUE"
			Case vbCancel	iRetVal = XL_DISPATCH_CANCEL
			End Select
		else	'Use JEUtilities.exe
			oButtonList.reset
			oButtonList.AddButton sRUN
			oButtonList.AddButton sSKIP
			oButtonList.AddButton sEND
			oButtonList.Caption = "Paused at row"
			oButtonList.Description = oRow.Cells(2, XL_KEYWORD).Value
			oButtonList.show
			sSelectedButton = oButtonList.SelectedButton
			vbButtonPressed = oButtonList.ButtonPressed
			Do while sSelectedButton = ""
				iYesNo = msgbox("You did not select an option." & Chr(13) & "Press Yes to try again or No to stop the test.", vbYesNo, "Try again.")
				if iYesNo = vbNo then
					exit do
				end if
				oButtonList.show
				sSelectedButton = oButtonList.SelectedButton
				vbButtonPressed = oButtonList.ButtonPressed
			loop

			Select Case sSelectedButton
				Case sRUN		iRetVal = XL_DISPATCH_SKIP
				Case sSKIP		oRow.Cells(2, XL_DISABLE).Value = "TRUE"
				Case sEND		iRetVal = XL_DISPATCH_CANCEL
			End Select

			Set oButtonList = Nothing
		end if

		' Set return value
		PauseRow = iRetVal
		oVBSFramework.oTraceLog.Exited(sRoutine & ", Keyword=""" & CStr(oRow.Cells(1, XL_KEYWORD).Value) & """")
	End Function


public Function temp(oRow)
	dim oPopUp
	
	set oPopUp = new clsPopUp
	
	oPopUp.ShowMsg "Popup for 20 seconds. If you press OK here I will close down straight away.", 20

	msgbox "Paused while popup is on screen.  If you click OK here the popup will close down. "
	
	oPopUp.CloseMsg
	
	temp = XL_DISPATCH_PASS

end function


'--------------------------------------------------------------------------------------------------
' Class End
'--------------------------------------------------------------------------------------------------
End Class

'--------------------------------------------------------------------------------------------------
' File execute code
'--------------------------------------------------------------------------------------------------
Public oUtility
Set oUtility = new clsUtility

oVBSFramework.oTestObjects.Add "UTILITY", oUtility
