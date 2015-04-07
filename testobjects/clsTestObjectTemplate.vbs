Option Explicit
''==================================================================================================
' clsTestObjectTemplate.vbs 								'*** PLEASE CHANGE THIS LINE IF YOU USE THIS TEMPLATE AS THE BASIS FOR YOUR OWN TEST OBJECT
'
' Purpose: 
'   *** PLEASE CHANGE THE TEXT BELOW IF YOU USE THIS TEMPLATE AS THE BASIS FOR YOUR OWN TEST OBJECT
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
Class clsTestObjectTemplate 									'*** PLEASE CHANGE THIS LINE IF YOU USE THIS TEMPLATE AS THE BASIS FOR YOUR OWN TEST OBJECT
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

		Const sRoutine = "clsTestObjectTemplate.RowDispatch"	'*** PLEASE CHANGE THIS LINE IF YOU USE THIS TEMPLATE AS THE BASIS FOR YOUR OWN TEST OBJECT
		oVBSFramework.oTraceLog.Entered(sRoutine & ", Keyword=""" & CStr(oRow.Cells(1, XL_KEYWORD).Value) & """")
		
        'Get keyword from spreadsheet row as a string
        sKeyword= CStr(oRow.Cells(1, XL_KEYWORD).Value)
        'Get the action string after the dot
        sKeyword = right(sKeyword, len(sKeyword)-inStrRev(sKeyword, "."))

		Select Case uCase(sAction)
			'Dispatch known keyword functions
			Case "ACTIONTEMPLATE"								'*** PLEASE CHANGE THIS LINE IF YOU USE THIS TEMPLATE AS THE BASIS FOR YOUR OWN TEST OBJECT
				iRetVal = ActionTemplate(oRow)					'*** PLEASE CHANGE THIS LINE IF YOU USE THIS TEMPLATE AS THE BASIS FOR YOUR OWN TEST OBJECT
			
			Case Else
				iRetVal = XL_DISPATCH_UNKNOWN 	' Unknown keyword
		End Select

		RowDispatch = iRetVal
		oVBSFramework.oTraceLog.Exited(sRoutine & ", Keyword=""" & CStr(oRow.Cells(1, XL_KEYWORD).Value) & """")
	End Function

	'==============================================================================================
	' Function/Sub: ActionTemplate(oRow)
	' Purpose:
	'
	' Parameters:
	'
	' Returns:
	'==============================================================================================
	Private Function ActionTemplate(oRow)						'*** PLEASE CHANGE THIS LINE IF YOU USE THIS TEMPLATE AS THE BASIS FOR YOUR OWN TEST OBJECT
		Const sRoutine = "clsTestObjectTemplate.ActionTemplate" '*** PLEASE CHANGE THIS LINE IF YOU USE THIS TEMPLATE AS THE BASIS FOR YOUR OWN TEST OBJECT
		oVBSFramework.oTraceLog.Entered(sRoutine)

		ActionTemplate = XL_DISPATCH_PASS						'*** PLEASE CHANGE THIS LINE IF YOU USE THIS TEMPLATE AS THE BASIS FOR YOUR OWN TEST OBJECT
		oVBSFramework.oTraceLog.Exited(sRoutine)
	End Function

'--------------------------------------------------------------------------------------------------
' Class End clsTestObjectTemplate
'--------------------------------------------------------------------------------------------------
End Class

'Registration code 
Public oTestObjectTemplate										'*** PLEASE CHANGE THIS LINE IF YOU USE THIS TEMPLATE AS THE BASIS FOR YOUR OWN TEST OBJECT
Set oTestObjectTemplate = New clsTestObjectTemplate 			'*** PLEASE CHANGE THIS LINE IF YOU USE THIS TEMPLATE AS THE BASIS FOR YOUR OWN TEST OBJECT
oVBSFramework.oTestObjects.Add "TestObjectTemplate", oTestObjectTemplate	'*** PLEASE CHANGE THIS LINE IF YOU USE THIS TEMPLATE AS THE BASIS FOR YOUR OWN TEST OBJECT
