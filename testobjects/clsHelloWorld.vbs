Option Explicit
''==================================================================================================
' clsHelloWorld.vbs 								'
'
' Purpose: 
'   
'
'	This file is a template for a Test Object that can be used to develop a new Test Object. 
'	It contains an example public RowDispatch method, an example private ActionTemplate method and 
'	an example of the framework registration code needed to integrate with the QTP Automation 
'	Framework.
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
Class clsHelloWorld							
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

		Const sRoutine = "clsHelloWorld.RowDispatch"
			
		oVBSFramework.oTraceLog.Entered(sRoutine & ", Keyword=""" & CStr(oRow.Cells(1, XL_KEYWORD).Value) & """")
		
        'Get keyword from spreadsheet row as a string
        sKeyword= CStr(oRow.Cells(1, XL_KEYWORD).Value)
        'Get the action string after the dot
        sKeyword = right(sKeyword, len(sKeyword)-inStrRev(sKeyword, "."))		
		
		Select Case uCase(sKeyword)
			'Dispatch known keyword functions		
			Case "SAYHELLOWORLD"							
				iRetVal = sayHelloWorldAction(oRow,sOutputParams,sLog)							
			Case Else
				iRetVal = XL_DISPATCH_UNKNOWN 	' Unknown keyword
		End Select

		RowDispatch = iRetVal
		oVBSFramework.oTraceLog.Exited(sRoutine & ", Keyword=""" & CStr(oRow.Cells(1, XL_KEYWORD).Value) & """")
	End Function

	'==============================================================================================
	' Function/Sub: sayHelloWorldAction(oRow,sOutputParams,sLog)
	' Purpose:
	'
	' Parameters:
	'
	' Returns:
	'==============================================================================================
	Private Function sayHelloWorldAction(oRow,sOutputParams,sLog)					
		Dim sName,iRetVal
		Const sRoutine = "clsHelloWorld.sayHelloWorldAction" 
		oVBSFramework.oTraceLog.Entered(sRoutine)
		
		sName = CStr(oRow.Cells(1, XL_PARM_001).Value)		
		
		iRetVal = sayHelloWorld(sName)			
		
		sOutputParams = "Hello=World" 
		sLog = sLog & "M: Hello World"
		sayHelloWorldAction = iRetVal	
						
		oVBSFramework.oTraceLog.Exited(sRoutine)
	End Function
	
	'==============================================================================================
	' Function/Sub: sayHelloWorldr(sName)
	' Purpose:
	'
	' Parameters:
	'
	' Returns: Creation Time value
	'==============================================================================================
	Private Function sayHelloWorld(sName)					
		Const sRoutine = "clsHelloWorld.sayHelloWorld" 
		oVBSFramework.oTraceLog.Entered(sRoutine)
		If sName = "" Then
			MsgBox "Hello World!"
		Else
			MsgBox "Hello " & sName & "!"
		End If
		
		sayHelloWorld = XL_DISPATCH_PASS	
		oVBSFramework.oTraceLog.Exited(sRoutine)
	End Function	
'--------------------------------------------------------------------------------------------------
' Class End clsHelloWorld
'--------------------------------------------------------------------------------------------------
End Class

'Registration code 
Public oHelloWorld										
Set oHelloWorld = New clsHelloWorld		
oVBSFramework.oTestObjects.Add "HELLOWORLD", oHelloWorld
