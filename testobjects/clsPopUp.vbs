Option Explicit
'==================================================================================================
' clsPopUp.vbs
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

Class clsPopUp

	'==============================================================================================
	' CLASS VARIABLES
	'==============================================================================================

	'----------------------------------------------------------------------------------------------
	' Public Variables
	'----------------------------------------------------------------------------------------------

	'----------------------------------------------------------------------------------------------
	' Private Variables
	'----------------------------------------------------------------------------------------------

	Private oWshShell
	Private oExec
	Private oPopUp

	'----------------------------------------------------------------------------------------------
	' CLASS_INITIALIZE
	'----------------------------------------------------------------------------------------------
	Private Sub Class_Initialize

	End Sub

	'==============================================================================================
	' Function/Sub:
	' Purpose:
	'==============================================================================================
	Public Sub ShowMsg(sMessage, iWait)
		Dim oFS, oFile
		'Dim oPopUp
		
		const FORWRITING =  2

		const sRoutine = "clsPopUp.ShowMsg"
		oVBSFramework.oTraceLog.Entered(sRoutine)

		on error resume next
		set oPopUp = createobject("JEUtilities.Popup")
		if oPopUp is Nothing then 'Use the fallback method.
			if IsObject(oWshShell) then
				CloseMsg
			end if

			Set oFS = CreateObject("Scripting.FileSystemObject")

			'Write the file
			set oFile = oFS.OpenTextFile("c:\PopUpMessage.vbs", FORWRITING, True)
			oFile.WriteLine "Dim sh, r"
			oFile.WriteLine "Set sh = createobject(""wscript.shell"")"
			oFile.WriteLine "r = sh.popup(""" & sMessage & """," & iWait & ",""Automation Information"")"
			oFile.WriteLine "set sh = nothing"
			oFile.Close

			Set oWshShell = CreateObject("WScript.Shell")
			Set oExec = oWshShell.Exec("wscript c:\PopUpMessage.vbs")
		else
			oPopUp.WindowTitle = "Progress information"
			oPopUp.ShowProgressBar = True
			oPopUp.Timeout = iWait
			oPopUp.Message = sMessage
			oPopUp.ShowMsg sMessage		
		end if
		on error goto 0

		oVBSFramework.oTraceLog.Exited(sRoutine)
	End Sub

	'==============================================================================================
	' Function/Sub:
	' Purpose:
	'==============================================================================================
	Public Sub CloseMsg()

		Dim oFile

		const sRoutine = "clsPopUp.CloseMsg"
		oVBSFramework.oTraceLog.Entered(sRoutine)

		on error resume next
		if oPopUp is Nothing then
			if IsObject(oWshShell) then
				oExec.Terminate
				Set oExec = nothing
				Set oWshShell = nothing
			end if
		else
			oPopUp.CloseMsg
		end if 
		on error goto 0

		oVBSFramework.oTraceLog.Exited(sRoutine)
	End Sub

'--------------------------------------------------------------------------------------------------
' Class End clsPopUp
'--------------------------------------------------------------------------------------------------
end class

'set oPopUp = new clsPopUp
'oPopUp.ShowMsg "Testing PopUp - 1", 0
'
'wait 5
'msgbox "Next"
'
'oPopUp.ShowMsg "Testing PopUp - 2", 2
'
'wait 5
'msgbox "Next"
'
'oPopUp.CloseMsg

