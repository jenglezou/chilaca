Option Explicit
''==================================================================================================
' clsSeleniumChrome.vbs 								'
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
Class clsSeleniumChrome							
	'==============================================================================================
	' CLASS VARIABLES
	'==============================================================================================

	'----------------------------------------------------------------------------------------------
	' Public Variables
	'----------------------------------------------------------------------------------------------
	
	'----------------------------------------------------------------------------------------------
	' Private Variables
	'----------------------------------------------------------------------------------------------
	Private oDriver

	'----------------------------------------------------------------------------------------------
	' CLASS_INITIALIZE
	'----------------------------------------------------------------------------------------------

	'==============================================================================================
	' CLASS SUBS & FUNCTIONS
	'==============================================================================================
	'==============================================================================================
	' Function/Sub: RowDispatch
	' Purpose:		RowDispatch (must be a public function) is what the QTP framework uses to 
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

		Const sRoutine = "clsSeleniumChrome.RowDispatch"
			
		oVBSFramework.oTraceLog.Entered(sRoutine & ", Keyword=""" & CStr(oRow.Cells(1, XL_KEYWORD).Value) & """")
		
        'Get keyword from spreadsheet row as a string
        sKeyword= CStr(oRow.Cells(1, XL_KEYWORD).Value)
        'Get the action string after the dot
        sKeyword = right(sKeyword, len(sKeyword)-inStrRev(sKeyword, "."))		
		
		Select Case uCase(sKeyword)
			'Dispatch known keyword functions		
			Case "START"							
				iRetVal = Start(oRow,sOutputParams,sLog)							
			Case "TEST"							
				iRetVal = Test(oRow,sOutputParams,sLog)							
			Case "STOP"							
				iRetVal = SeleniumStop(oRow,sOutputParams,sLog)							
			Case Else
				iRetVal = XL_DISPATCH_UNKNOWN 	' Unknown keyword
		End Select

		RowDispatch = iRetVal
		oVBSFramework.oTraceLog.Exited(sRoutine & ", Keyword=""" & CStr(oRow.Cells(1, XL_KEYWORD).Value) & """")
	End Function

	'==============================================================================================
	' Function/Sub: Start(oRow,sOutputParams,sLog)
	' Purpose:
	'
	' Parameters:
	'
	' Returns:
	'==============================================================================================
	Private Function Start(oRow,sOutputParams,sLog)					
		Dim iRetVal
		Const sRoutine = "clsSeleniumChrome.Start" 
		oVBSFramework.oTraceLog.Entered(sRoutine)
		
		iRetVal = XL_DISPATCH_PASS
        Set oDriver = CreateObject("Selenium.ChromeDriver")
		
		Start = iRetVal	
						
		oVBSFramework.oTraceLog.Exited(sRoutine)
	End Function
	
	'==============================================================================================
	' Function/Sub: Test(oRow,sOutputParams,sLog)
	' Purpose:
	'
	' Parameters:
	'
	' Returns:
	'==============================================================================================
	Private Function Test(oRow,sOutputParams,sLog)					
		Dim sName,iRetVal
		Const sRoutine = "clsSeleniumChrome.Test" 
		oVBSFramework.oTraceLog.Entered(sRoutine)
		
		iRetVal = XL_DISPATCH_PASS
		sName = CStr(oRow.Cells(1, XL_PARM_001).Value)		
		
        oDriver.Get "https://www.google.co.uk"
        oDriver.FindElementByName("q").SendKeys "Eiffel tower" & vbLf
        msgbox "Title=" & oDriver.Title & vbLF & "Click OK to terminate"
        msgbox "Title=" & oDriver.Title & vbLF & "Click OK to terminate"
		
		Test = iRetVal	
						
		oVBSFramework.oTraceLog.Exited(sRoutine)
	End Function

	'==============================================================================================
	' Function/Sub: SeleniumStop(oRow,sOutputParams,sLog)
	' Purpose:
	'
	' Parameters:
	'
	' Returns:
	'==============================================================================================
	Private Function SeleniumStop(oRow,sOutputParams,sLog)					
		Dim iRetVal
		Const sRoutine = "clsSeleniumChrome.SeleniumStop" 
		oVBSFramework.oTraceLog.Entered(sRoutine)
		
		iRetVal = XL_DISPATCH_PASS
		oDriver.Quit
		
		SeleniumStop = iRetVal	
						
		oVBSFramework.oTraceLog.Exited(sRoutine)
	End Function

'--------------------------------------------------------------------------------------------------
' Class End clsSeleniumChrome
'--------------------------------------------------------------------------------------------------
End Class

'Registration code 
Public oSeleniumChrome										
Set oSeleniumChrome = New clsSeleniumChrome		
oVBSFramework.oTestObjects.Add "SELENIUM", oSeleniumChrome
