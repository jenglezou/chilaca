Option Explicit
''==================================================================================================
' clsPassFailContinue.vbs
'
' Purpose
'
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
Class clsPassFailContinue
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

	'==============================================================================================
	' CLASS SUBS & FUNCTIONS
	'==============================================================================================
	'==============================================================================================
	' Function/Sub:
	' Purpose:
	'
	' Parameters:
	'
	' Returns:
	'==============================================================================================
	Public Function RowDispatch(oRow, sLog, sOutputParams)
		Dim iRetVal, sKeyword, hWin, sWindowTitle

		Const sRoutine = "clsPassFailContinue.RowDispatch"
		oVBSFramework.oTraceLog.Entered(sRoutine & ", Keyword=""" & CStr(oRow.Cells(1, XL_KEYWORD).Value) & """")

		' Get keyword from spreadsheet as string
		sKeyword = CStr(oRow.Cells(1, XL_KEYWORD).Value)
		sKeyword = right(sKeyword, len(sKeyword)-inStrRev(sKeyword, "."))

		sLog = sLog & "D added this in clsPassFailContinue. Action: " & sKeyword & vblf

		Select Case uCase(sKeyword)
			'Dispatch known keyword functions
			'Case "<keyword>"				iRetVal = K_<Keyword>(oRow)
			Case "PASS"
				iRetVal = XL_DISPATCH_PASS
			Case "FAIL"
				iRetVal = XL_DISPATCH_FAIL
			Case "FAILCONTINUE"
				iRetVal = XL_DISPATCH_FAILCONTINUE
			Case "UNEXPECTEDERROR"
				Window("name:=shouldgiveanerror").Close
				iRetVal = XL_DISPATCH_PASS
			Case "FILEERROR"
				iRetVal = FileError(oRow)
			Case "OUTPUT"
				iRetVal = OutputAction(oRow, sLog, sOutputParams)
			Case Else
				iRetVal = XL_DISPATCH_UNKNOWN 	' Unknown keyword
		End Select

		RowDispatch = iRetVal
		oVBSFramework.oTraceLog.Exited(sRoutine & ", Keyword=""" & CStr(oRow.Cells(1, XL_KEYWORD).Value) & """")
	End Function

	'==============================================================================================
	' Function/Sub:
	' 
	'==============================================================================================
	Private Function FileError(oRow)
		Dim iRetval, sTextFile
		Dim oFS, oFile
		
		Set oFS = CreateObject("Scripting.FileSystemObject")

		'Create the textfile names
		'sTextFile = oFS.GetParentFolderName("C:\temp\notexist.xxx") & "\" &  oFS.GetBaseName("C:\temp\notexist.xxx") & ".txt"
		Set oFile = oFS.OpenTextFile("C:\temp\notexist.xxx", 1, False)


'msgbox sTextFile
		Set oFS = Nothing		
			
		FileError = XL_DISPATCH_PASS
	End Function
	
	'==============================================================================================
	' Function/Sub:
	' 
	'==============================================================================================
	Private Function OutputAction(oRow, sLog, sOutputParams)
		Dim sTempCellValue
		
		sTempCellValue = Trim(CStr(oRow.Cells(1, XL_PARM_001).Value))
		If sTempCellValue <> "" Then sOutputParams = sOutputParams & "PARAM1=" & sTempCellValue & ","
		
		sTempCellValue = Trim(CStr(oRow.Cells(1, XL_PARM_002).Value))
		If sTempCellValue <> "" Then sOutputParams = sOutputParams & "PARAM2=" & sTempCellValue & ","
		
		sTempCellValue = Trim(CStr(oRow.Cells(1, XL_PARM_003).Value))
		If sTempCellValue <> "" Then sOutputParams = sOutputParams & "PARAM3=" & sTempCellValue & ","
		
		sTempCellValue = Trim(CStr(oRow.Cells(1, XL_PARM_004).Value))
		If sTempCellValue <> "" Then sOutputParams = sOutputParams & "PARAM4=" & sTempCellValue & ","
			
		OutputAction = XL_DISPATCH_PASS
	End Function
'--------------------------------------------------------------------------------------------------
' Class End clsPassFailContinue
'--------------------------------------------------------------------------------------------------
end class

Private oPassFailContinue
Set oPassFailContinue = new clsPassFailContinue
oVBSFramework.oTestObjects.Add "PassFailContinue", oPassFailContinue
