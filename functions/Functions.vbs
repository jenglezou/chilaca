Option Explicit
'==================================================================================================
'ROBODOC header blocks follow ...
'****L* Function Libraries/Functions
'  SYNOPSIS
'	Function Library - Functions
'  OVERVIEW
' 	The VBSCript module containing the Functions function library
'	All miscellaneous support functions and subs that are used by the application object classes go here.
'	May have to split this into smaller files or classes if a significant number of functional groups 
'	become apparent. In general, the functions and subs here will be Public.
'*****
'==================================================================================================

Public sRunAsUser 'REM why is it here?
'Public bPerformance		' Flag to add performance csv file
'Public sPerformanceFile	'Performance file
'==============================================================================================
Public Function GetOpenWindowTitles()
	Dim oDesc, oObjects, oObj
	Dim iCounter
	ReDim arrResult(-1)
	
	Set oDesc = Description.Create()
	
	oDesc("micclass").Value = "Window"
	
	Set oObjects = Desktop.ChildObjects(oDesc)
	
	For iCounter = 0 To oObjects.Count - 1
		Set oObj = oObjects(iCounter)
		
		ReDim Preserve arrResult(UBound(arrResult) + 1)
		arrResult(UBound(arrResult)) = oObj.GetROProperty("title")
		
		Set oObj = Nothing
	Next
	
	GetOpenWindowTitles = arrResult
End Function

'==============================================================================================
Public Function ArrayContainsItem(arrItems, sItemtoFind)
	Dim sItem
	
	ArrayContainsItem = False
	
	For Each sItem in arrItems
		If sItem = sItemToFind Then
			ArrayContainsItem = True
			Exit For
		End If
	Next
End Function

'==============================================================================================
Public Function EscapeSpecialCharsWithBackslash(sString)
	EscapeSpecialCharsWithBackslash = Replace( Replace( Replace( Replace(sString, ".", "\."), "(", "\("), ")", "\)"), "+", "\+")
End Function

'==============================================================================================
'ROBODOC header blocks follow ...
'****f* Functions/EvalWindow,EvalWindow{Functions}
'  SYNOPSIS
'	Public Function EvalWindow(arrWindowDescription, iWaitTime)
'  OVERVIEW
' 	
'*****
'==============================================================================================
Public Function EvalWindow(arrWindowDescription, iWaitTime)
	Dim sTempString
	Dim sWindowFunc
	Dim i

	if not isarray(arrWindowDescription) then 
		sTempString = arrWindowDescription
		redim arrWindowDescription(0)
		arrWindowDescription(0) = sTempString
	End if

	sWindowFunc = "Window("
	For i = 0 to ubound(arrWindowDescription)
		sWindowFunc = sWindowFunc & """" & arrWindowDescription(i) & """" & ","
	Next
	sWindowFunc = left(sWindowFunc,len(sWindowFunc) -1) & ").Exist(0)"

	Do While (eval(sWindowFunc) = False) And (iWaitTime > 0)
		iWaitTime = iWaitTime - 1
		'WaitTime(1)
		WaitTime(1)
	Loop
	if iWaitTime > 0 Then 
		EvalWindow = True
	Else
		EvalWindow = False
	End If
End Function

'==============================================================================================
'ROBODOC header blocks follow ...
'****f* Functions/ExistWindowDoubleCheck,ExistWindowDoubleCheck{Functions}
'  SYNOPSIS
'	Public Function ExistWindowDoubleCheck(arrWindowDescription, iWaitTime)
'  OVERVIEW
' 	
'*****
'==============================================================================================
Public Function ExistWindowDoubleCheck(arrWindowDescription, iWaitTime)
	Dim sTempString
	Dim sWindowFunc
	Dim i
	
	if not isarray(arrWindowDescription) then 
		sTempString = arrWindowDescription
		redim arrWindowDescription(0)
		arrWindowDescription(0) = sTempString
	End if

	sWindowFunc = "Window("
	For i = 0 to ubound(arrWindowDescription)
		sWindowFunc = sWindowFunc & """" & arrWindowDescription(i) & """" & ","
	Next
	sWindowFunc = Left(sWindowFunc,len(sWindowFunc) -1) & ").Exist(" & iWaitTime & ")"
	
	ExistWindowDoubleCheck = False
	
	If Eval(sWindowFunc) = True Then
		If Eval(sWindowFunc) = True Then
			ExistWindowDoubleCheck = True
		End If
	End If
End Function

'TODO: find more general way
'==============================================================================================
' Function/Sub:
'==============================================================================================
'REM QTP specific method 
Public Function DesktopObjectCount(sClass, sDescription)
	Dim oDescription, dictDescription, arrKeys, arrItems, sTempString, i
	
	Set dictDescription = CreateObject("scripting.dictionary")
	
	SimpleDictKeyValuePairs dictDescription, sDescription, ",", "="

	arrKeys = dictDescription.Keys
	arrItems = dictDescription.Items
	
	Set oDescription = Description.Create()
	oDescription("micclass").Value = sClass
	
	For i = 0 to dictDescription.Count - 1
		oDescription(arrKeys(i)).Value = arrItems(i)
	Next
	
	DesktopObjectCount = Desktop.ChildObjects(oDescription).Count
	
	Set dictDescription = Nothing
	Set oDescription = Nothing
End Function

'==============================================================================================
'ROBODOC header blocks follow ...
'****f* Functions/Screenshot,Screenshot{Functions}
'  SYNOPSIS
'	Public Function Screenshot(sFile, bUploadToQCRun)
'  OVERVIEW
' 	
'*****
'==============================================================================================
 Public Function Screenshot(sFile, bUploadToQCRun)
	Dim oTAScreenshot 
	Dim bRetVal 
	Dim sLog
	
	sLog = ""	
	Set oTAScreenshot = CreateObject("TAUtility.Screenshot")
	WaitTime(200)
	bRetVal = oTAScreenshot.MakeScreenshot(sFile, sLog)
	WaitTime(200)
	
	If Not bRetVal Then
		oVBSFramework.oTraceLog.Message "Screenshot not made", "LOG: " & sLog, LOG_ERROR
		oVBSFramework.oTraceLog.StepMessage "Screenshot not made", XL_DISPACH_FAIL, sLog, Null, ""
	Else
		If bUploadToQCRun = True Then		 
			If oQC.IsQCRun() Then
				oQC.UploadAttachmentToQCRun(sFile)
				oVBSFramework.oTraceLog.Message "Screenshot loaded to ALM. File name: " & sFile, LOG_MESSAGE
			End If
		End If
		oVBSFramework.oTraceLog.Message "Screenshot made. File name: " & sFile, LOG_MESSAGE
	End If
	Set oTAScreenshot = Nothing
 End Function

'==============================================================================================
'ROBODOC header blocks follow ...
'***if* Functions/ArrayDifferenceNewItems,ArrayDifferenceNewItems{Functions}
'  SYNOPSIS
'	Private Function ArrayDifferenceNewItems(arr1, arr2, arrNewItems)
'  OVERVIEW
'	arr1 - first array (of Strings)
'	arr2 - second array (of Strings)
'	arrNewItems - resulting array (of Strings) contain Strings from arr2 which are not in arr1
'	arrNewItems has to be dynamic array (ReDim)
'*****
'==============================================================================================
Private Function ArrayDifferenceNewItems(arr1, arr2, arrNewItems)
	Dim sItemArr1, sItemArr2
	Dim bWindowExisted
	Dim iCounterNewWindows

	ArrayDifferenceNewItems = False

	iCounterNewWindows = 0
	
	ReDim arrNewItems(0)
	
	arrNewItems(0) = Empty

	For Each sItemArr2 In arr2
		bWindowExisted = false
		For Each sItemArr1 In arr1
			If sItemArr1 = sItemArr2 Then
				bWindowExisted = True
				Exit For
			End If
		Next
		If not bWindowExisted Then
			ArrayDifferenceNewItems = True
			iCounterNewWindows = iCounterNewWindows + 1
			ReDim Preserve arrNewItems(iCounterNewWindows - 1)
			arrNewItems(iCounterNewWindows - 1) = sItemArr2
		End If
	Next
End Function
	
'==============================================================================================
'ROBODOC header blocks follow ...
'****f* Functions/DictKeyValuePairs,DictKeyValuePairs{Functions}
'  SYNOPSIS
'	Public Function DictKeyValuePairs(oDictData, sInput, sPairInitiator, sPairTerminator, sKeyValueSeparator)
'  OVERVIEW
'	Takes an XML-style input string of the form <key1=value1> <key2=value2> and 
'	populates the supplied dictionary object with the keys and values.
'	Actually, the start/end delimiters for the pairs and the key/value separator are configurable.
'  EXAMPLE
'	Test: should return "0:a is b, 1:c is , 2:x is , 0:a is 1, 1:b is 7, "
'	set d = CreateObject("Scripting.Dictionary")
'	DictKeyValuePairs d, "<<<<<>>>><a= b><c =><c=d>><=f><x><<>>>", "<", ">", "="
'	arrItems = d.Items
'	arrkeys = d.keys
'	For i = 0 to d.Count-1
'		sMessage = sMessage & i & ":" & arrkeys(i) & " is " & arritems(i) & ", "
'	Next
'
'	DictKeyValuePairs d, "[a,1[b,7]]", "[", "]", ","
'	arrItems = d.Items
'	arrkeys = d.keys
'	For i = 0 to d.Count-1
'		sMessage = sMessage & i & ":" & arrkeys(i) & " is " & arritems(i)  & ", "
'	Next
'	msgbox smessage
'*****
'==============================================================================================
Public Function DictKeyValuePairs(oDictData, sInput, sPairInitiator, sPairTerminator, sKeyValueSeparator)
	Dim sInput1, sInput2
	Dim arrPairs, iPair
	Dim arrKeyValue, sKey, sValue

	'Split the inpput string into key value pairs
	arrPairs = Split(sInput, sPairInitiator)

	oDictData.RemoveAll
	For iPair = 0 to ubound(arrPairs)
		sKey = ""
		sValue = ""
		arrKeyValue = split(replace(arrPairs(iPair), sPairTerminator, ""), sKeyValueSeparator)
		If ubound(arrKeyValue) >= 0 then 
			sKey = Trim(arrKeyValue(0))
			If ubound(arrKeyValue) > 0 then sValue = Trim(arrKeyValue(1))
			If sKey <> "" And Not oDictData.Exists(sKey) Then
				oDictData.Add sKey, sValue
			End If
		end if
	Next

	DictKeyValuePairs = True
End Function

'==============================================================================================
'ROBODOC header blocks follow ...
'****f* Functions/SimpleDictKeyValuePairs,SimpleDictKeyValuePairs{Functions}
'  SYNOPSIS
'	Public Function SimpleDictKeyValuePairs(dictData, sString, sDelimiter, sSeparator)
'  OVERVIEW
' 	Populates a dictionary object from a list of delimited name value pairs
'	eg name1=value1,n2=v2,n3=v3 ...
'*****
'==============================================================================================
Public Function SimpleDictKeyValuePairs(dictData, sString, sDelimiter, sSeparator)
	Dim iRetVal
	Dim i, arrKeyValue

	iRetVal = 0
	arrKeyValue = Split(sString, sDelimiter)	'The key=value pairs in an array
	For i = 0 to UBound(arrKeyValue)
		if instr(arrKeyValue(i), sSeparator) > 0 Then
			dictData.add Trim(Split(Trim(arrKeyValue(i)),sSeparator)(0)),  Trim(Right(arrKeyValue(i), Len(arrKeyValue(i)) - Len(Split(arrKeyValue(i),sSeparator)(0))-1))'Trim(Split(Trim(arrKeyValue(i)),sSeparator)(1))
			iRetVal = i
		Else
			iRetVal = 0
			Exit for
		End if 
	Next

	SimpleDictKeyValuePairs = iRetVal
End Function

'==============================================================================================
' Function/Sub:
'==============================================================================================
Public Function DictKeyIndexValue(dictData, sKey, ByVal iIndex, sLog)
	Dim sResult, arrKeyValues

	sResult = ""
	iIndex = iIndex - 1
	
	If dictData.Exists(UCase(sKey)) Then
		arrKeyValues = Split(dictData.Item(UCase(sKey)),",")
		If iIndex => 0 And iIndex <= UBound(arrKeyValues) Then
			sResult = arrKeyValues(iIndex)
		Else
			sLog = sLog & "W: Array index out of bounds" & vbLf
		End If
	Else
		sLog = sLog & "W: Key value '" & sKey & "' is not present in dictionary object" & vbLf
	End If
	
	DictKeyIndexValue = sResult
End Function

'==============================================================================================
' Function/Sub:
'==============================================================================================
Public Function TextDataDelimited(oDictDelimitedData, sOutput, sInFile, iDataRow, iHeaderRow, sDelimiter)
	Dim oFS, oInFile
	Dim iRow, bFailed
	Dim sInLine
	Dim arrHeading, arrData

	Const ForReading = 1
	Const ForWriting = 2

	arrHeading = Array(1,2,3,4,5,6,7,8,9,10,11,12,13,14,15,16,17,18,19,20)	'Default array for key names in dictionary object 
	sOutput = ""
	bFailed = False

	Set oFS = CreateObject("Scripting.FileSystemObject")
	if  oFS.FileExists(sInFile)= True Then							'Only run if input file exists.
		Set oInFile = oFS.OpenTextFile(sInFile, ForReading, True)
		iRow = 1
		Do While Not oInFile.AtEndOfStream
			sInLine = oInFile.ReadLine
			If iHeaderRow = iRow Then arrHeading = split(sInLine, sDelimiter)	'Override  the default headings

			'Get the data into the dictionary object using the Header array as keys
			If iDataRow  = iRow Then
				arrData = split(sInLine, sDelimiter)
				For iField = 1 to UBound(arrData)
					oDictDelimitedData.add arrHeading(iField), iArrData(iField)
					sOutput = sOutput & iArrData(iField) & ","
				Next
				Exit do
			End if
		Loop

		oInFile.Close
		Set oInFile = Nothing'Tidy up objects.
	Else
		bFailed = True		'File not found
	End if

	Set oFS = Nothing
	TextDataDelimited = bFailed
End Function

'==============================================================================================
' Function/Sub: MonthLookup(sMonth)
'==============================================================================================
Public Function MonthLookup(sMonth)
	Dim iRetval

	Select case UCase(mid(sMonth,1,3))
	Case "JAN"
		iRetval = 1	
	Case "FEB"
		iRetval = 2	
	Case "MAR"
		iRetval = 3	
	Case "APR"
		iRetval = 4	
	Case "MAY"
		iRetval = 5	
	Case "JUN"
		iRetval = 6	
	Case "JUL"
		iRetval = 7	
	Case "AUG"
		iRetval = 8	
	Case "SEP"
		iRetval = 9	
	Case "OCT"
		iRetval = 10	
	Case "NOV"
		iRetval = 11	
	Case "DEC"
		iRetval = 12	
	Case Else
		iRetval = 0	
	End Select

	MonthLookup = iRetval

End Function

'==============================================================================================
' Function/Sub: YYYYMMDDHHMMSS(dDate)
'==============================================================================================
Public Function YYYYMMDDHHMMSS(dDate)
	YYYYMMDDHHMMSS = Year(dDate) & Right("0" & Month(dDate), 2) & Right("0" & Day(dDate), 2) _
					& Right("0" & Hour(dDate), 2) & Right("0" & Minute(dDate), 2) &  Right("0" & Second(dDate), 2)
End Function

'==============================================================================================
' Function/Sub: YYYYMMDDHHMM(dDate)
'==============================================================================================
Public Function YYYYMMDDHHMM(dDate)
	YYYYMMDDHHMM = Year(dDate) & Right("0" & Month(dDate), 2) & Right("0" & Day(dDate), 2) _
					& Right("0" & Hour(dDate), 2) & Right("0" & Minute(dDate), 2)
End Function

'==============================================================================================
' Function/Sub: YYYYMMDD(dDate)
'==============================================================================================
Public Function YYYYMMDD(dDate)
	YYYYMMDD = Year(dDate) & Right("0" & Month(dDate), 2) & Right("0" & Day(dDate), 2)
End Function

'==============================================================================================
' Function/Sub: LatestFilenameByDate(sFoldername, sFileType)
' Purpose: The latest file matching the filetype in the folder.
'==============================================================================================
Public Function LatestFilenameByDate(sFoldername, sFileType)
	Dim oFS, oFolder, fc, oFile
	Dim sLatestFilename, sDateLastModified

	'on Error Resume Next

	Set oFS   = CreateObject("Scripting.FileSystemObject")
	Set oFolder = oFS.GetFolder(sFoldername)
	'if err = 0 Then
		Set fc = oFolder.files
		sLatestFilename = ""
		sDateLastModified = "01/01/1900"
		For Each oFile in fc
			If(instr(len(oFile)-len(sFileType)+ 1, ucase(oFile), ucase(sFileType),1)= len(oFile)-len(sFileType)+ 1)Then
				'Found matching file
				If CDate(oFile.datelastmodified) > CDate(sDateLastModified) Then
					sDateLastModified = oFile.datelastmodified
					sLatestFilename =  ucase(oFile)
				End if
			End if
		Next
	'End if 'err = 0

	Set oFile = Nothing
	Set oFolder = Nothing
	Set oFS = Nothing

	LatestFilenameByDate = sLatestFilename
End Function

'==============================================================================================
' Function/Sub: CountFiles(sFoldername)
' Purpose: How many files  in the folder.
'==============================================================================================
Public Function CountFiles(sFoldername)
	Dim oFS, oFolder, oFiles

	Set oFS = CreateObject("Scripting.FileSystemObject")
	CountFiles = oFS.GetFolder(sFoldername).Files.Count
	Set oFS = Nothing
End Function

'==============================================================================================
' Function/Sub: CountFiles(sFoldername)
' Purpose: How many lines in text file.
'==============================================================================================
Public Function CountFileLines(sFilename)
	Dim oFS, oFile
	
	Set oFS = CreateObject("Scripting.FileSystemObject")
	Set oFile = oFS.OpenTextFile(sFilename)	
	
	CountFileLines = Ubound(Split(oFile.ReadAll, vbNewLine))' - 1 'REM put out of here
	oFile.Close
	Set oFile = Nothing
	Set oFS = Nothing
End Function


'==============================================================================================
' Function/Sub: MoveFiles(sSourceFolder, sDestinationFolder)
' Purpose: 
'==============================================================================================
Public Sub MoveFiles(sSourceFolder, sDestinationFolder)
	Dim oFS

	Set oFS = CreateObject("Scripting.FileSystemObject")

	'Create the folder
	On Error resume next
	oFS.CreateFolder(sDestinationFolder)
	oFS.MoveFile sSourceFolder & "\*.*", sDestinationFolder
	On Error goto 0

	Set oFS = Nothing
End Sub

'==============================================================================================
' Function/Sub: CopyFiles(sSourceFolder, sDestinationFolder)
' Purpose: 
'==============================================================================================
Public Sub CopyFiles(sSourceFolder, sDestinationFolder)
	Dim oFS

	Set oFS = CreateObject("Scripting.FileSystemObject")

	'Create the folder
	On Error resume next
	oFS.CreateFolder(sDestinationFolder)
	oFS.CopyFile sSourceFolder & "\*.*", sDestinationFolder
	On Error goto 0

	Set oFS = Nothing
End Sub

'==============================================================================================
' Function/Sub: DeleteFiles(sSourceFolder)
' Purpose: 
'==============================================================================================
Public Sub DeleteFiles(sSourceFolder)
	Dim oFS

	Set oFS = CreateObject("Scripting.FileSystemObject")

	On Error resume next
	oFS.DeleteFile sSourceFolder & "\*.*"
	On Error goto 0

	Set oFS = Nothing
End Sub

'==============================================================================================
' Function/Sub: LatestFileDate(sFoldername, sFileType)
' Purpose: The latest file matching the filetype in the folder.
'==============================================================================================
Public Function LatestFileDate(sFoldername, sFileType)
	Dim oFS, oFolder, oFC, oFile
	Dim sLatestDate

	'on Error Resume Next

	Set oFS   = CreateObject("Scripting.FileSystemObject")
	Set oFolder = oFS.GetFolder(sFoldername)
	Set oFC = oFolder.Files

	sLatestDate = "01/01/1900"
	For Each oFile in oFC
		If(instr(len(oFile)-len(sFileType)+ 1, ucase(oFile), ucase(sFileType),1)= len(oFile)-len(sFileType)+ 1)Then
			'Found matching file
			If CDate(oFile.datelastmodified) > CDate(sLatestDate) Then
				sLatestDate = oFile.datelastmodified
			End if
		End if
	Next

	Set oFile = Nothing
	Set oFC = Nothing
	Set oFolder = Nothing
	Set oFS = Nothing

	LatestFileDate = sLatestDate
End Function

'==============================================================================================
' Function/Sub: Encrypt(sString)
' Purpose: 
'==============================================================================================
Public Function Encrypt(sString)
	Dim sKey
	Dim iLenKey, iKeyPos, iLenStr, i, sNewStr

	sKey = "encryption"
	sNewStr = ""
	iLenKey = Len(sKey)
	iKeyPos = 1
	iLenStr = Len(sString)

	sString = StrReverse(sString)
	For i = 1 To iLenStr
		sNewStr = sNewStr & chr(asc(Mid(sString, i, 1)) + Asc(Mid(sKey, iKeyPos, 1)))
		iKeyPos = iKeyPos + 1
		If iKeyPos > iLenKey Then iKeyPos = 1
	Next
	encrypt = sNewStr
End Function

'==============================================================================================
' Function/Sub: Decrypt(sString)
' Purpose: 
'==============================================================================================
Public Function Decrypt(sString)
	Dim sKey
	Dim iLenKey, iKeyPos, iLenStr, i, sNewStr

	sKey = "encryption"
	sNewStr = ""
	iLenKey = Len(sKey)
	iKeyPos = 1
	iLenStr = Len(sString)

	sString=StrReverse(sString)
	For i = iLenStr To 1 Step -1
		sNewStr = sNewStr & chr(asc(Mid(sString, i, 1)) - Asc(Mid(sKey, iKeyPos, 1)))
		iKeyPos = iKeyPos + 1
		If iKeyPos > iLenKey Then iKeyPos = 1
	Next
	sNewStr=StrReverse(sNewStr)
	Decrypt = sNewStr
End Function

'==============================================================================================
' Function/Sub: ClipboardText()
' Purpose: 
'==============================================================================================
Public Function ClipboardText()
	Dim oHtml
	
	Set oHtml = CreateObject("htmlfile")
	ClipboardText = oHtml.ParentWindow.ClipboardData.GetData("text")
	Set oHtml = nothing	
End Function

'==============================================================================================
' Function/Sub: GetQCWindowsUser()
' Purpose: 
'==============================================================================================
Public Function GetQCWindowsUser()
	Dim sUserName
	Dim oWshShell

	'Initialise the user as windows user 
	Set oWshShell = CreateObject("WScript.Shell")
	sUserName = oWshShell.ExpandEnvironmentStrings("%UserName%")
	Set oWshShell = Nothing
	
	'if test is run from QC, then use the QC username
	If QCUtil.IsConnected then 
		If Not QCUtil.CurrentTest Is Nothing Then
			sUserName = QCUtil.QCConnection.UserName
		End If
	End If

	If Trim(sRunAsUser) = "" Or UCase(Trim(sRunAsUser)) = "DEFAULT"  Then  
		GetQCWindowsUser = sUserName
	Else 
		GetQCWindowsUser = Trim(sRunAsUser)
	End If 
	
End Function

'==============================================================================================
' Function/Sub: SetRunAsUser()
' Purpose: 
'==============================================================================================
Public Sub SetRunAsUser(sUser)
	sRunAsUser = sUser
End Sub

'==============================================================================================
' Function/Sub: GetRunAsUser()
' Purpose: 
'==============================================================================================
Public Function GetRunAsUser()
	GetRunAsUser = sRunAsUser
End Function

'==============================================================================================
' Function/Sub: WritePerformanceFile()
' Purpose: 
'==============================================================================================
Public Sub WritePerformanceFile(sFile, sLine, iMode)
	Dim oFS, oFile
	Set oFS = CreateObject("Scripting.FileSystemObject")
			
	'Set oFile = oFS.OpenTextFile("C:\test.csv", ForAppending)
	If iMode = 0 Then
		Set oFile = oFS.OpenTextFile(sFile, 8, true)	'appending
		oFile.WriteLine sLine
		oFile.Close
	Else
'		SetPerformance True
		'sPerformanceFile = sFile
		If Not oFS.FileExists(sFile) Then
			Set oFile = oFS.OpenTextFile(sFile, 2, true)	'writing	
			oFile.WriteLine sLine
			oFile.Close
		End If
	End If	
End Sub

'==============================================================================================
' Function/Sub: WritePerformanceFileWithQCTestPath()
' Purpose: 
'==============================================================================================
Public Sub WritePerformanceFileWithQCTestPath(sFile, sLine, iMode)
	Dim sQCPath
	
	sQCPath = ""
	'if test is run from QC, then use the QC username
	If QCUtil.IsConnected then 
		If Not QCUtil.CurrentTest Is Nothing Then
			sQCPath = "," & QCUtil.CurrentTestSetTest.TestSet.TestSetFolder.Path
		End If
	End If
	WritePerformanceFile sFile, sLine & sQCPath, iMode
End Sub

'==============================================================================================
' Function/Sub: SetPerformance
' Purpose: 
'==============================================================================================
'Public Function SetPerformance(bBoolean)
'	bPerformance = bBoolean
'End Function


'==============================================================================================
' Function/Sub: IsPerformance()
' Purpose: 
'==============================================================================================
'Public Function IsPerformance()
'	IsPerformance = bPerformance
'End Function

'==============================================================================================
' Function/Sub: GetRunAsUser()
' Purpose: 
'==============================================================================================
'Public Function GetPerformanceFile()
'	GetPerformanceFile = sPerformanceFile
'End Function

'==============================================================================================
' Function/Sub: KeePassDBFile()
' Purpose: 
'==============================================================================================
Public Function KeePassDBFile(sFolder, sUserNameOverride)
	If Trim(sUserNameOverride) = "" Or UCase(Trim(sUserNameOverride)) = "DEFAULT"  Then  
		KeePassDBFile = sFolder & "\VBSFramework" & GetQCWindowsUser() & ".kdbx"
	Else
		KeePassDBFile = sFolder & "\VBSFramework" & sUserNameOverride & ".kdbx"
	End If
End Function

'==============================================================================================
' Function/Sub: KeePassKeyFile()
' Purpose: 
'==============================================================================================
Public Function KeePassKeyFile(sFolder, sUserNameOverride)
	Dim oFS, sKeyFile

	Set oFS = CreateObject("Scripting.FileSystemObject")
	If Trim(sUserNameOverride) = "" Or UCase(Trim(sUserNameOverride)) = "DEFAULT"  Then  
		sKeyFile = sFolder & "\KPKeyFile" & GetQCWindowsUser()
	Else
		sKeyFile = sFolder & "\KPKeyFile" & sUserNameOverride
	End If

	if Not oFS.FileExists(sKeyFile) then
		sKeyFile = sKeyFile & ".key"
	End If

	Set oFS = Nothing

	KeePassKeyFile = sKeyFile

End Function

'==============================================================================================
' Function/Sub: KeePassUserName()
' Purpose: 
'==============================================================================================
Public Function KeePassUserName(sKeePassFolder, sDBFile, sKeyFile, sTitle)
	Dim oWshShell, oExec, oClipboard',oSys3
	Dim iRetryCount
	Dim sRetVal

	'Set oSys3 = CreateObject("JSSys3.ops")
	'oSys3.SendTextCB("")
	'Set oClipboard = CreateObject("Mercury.Clipboard")
	'oClipboard.SetText ""
	Set oClipboard = CreateObject("TAUtility.TAClipboard")
	Set oWshShell = CreateObject("WScript.Shell")

	For iRetryCount = 1 to 5
		Set oExec = oWshShell.Exec(sKeePassFolder & "KPScript -c:clipusername " & """" & sDBFile & _
									""" -keyfile:""" & sKeyFile & """ -ref-Title:""" & sTitle & """")
		Do While oExec.Status = 0
			'WaitTime 0, 100
			WaitTime(100)
		Loop

'		sRetVal = ClipboardText()
'		oSys3.GetTextCB sRetVal
		'sRetVal = oClipboard.GetText()
		sRetVal = oClipboard.GetTextFromClipboard()
		'msgbox iRetryCount & ":" & sRetVal
		If Not (sRetVal = "" Or IsNull(sRetVal) Or IsEmpty(sRetVal)) Then
			oClipboard.ClearClipboard()
			Exit For
		End If
	Next
	
	'Set oSys3 = Nothing
	'Set oClipboard = Nothing
	Set oExec = Nothing
	Set oWshShell = Nothing

	KeePassUserName = sRetVal
End Function

'==============================================================================================
' Function/Sub: KeePassPassword()
' Purpose: 
'==============================================================================================
Public Function KeePassPassword(sKeePassFolder, sDBFile, sKeyFile, sTitle)
	Dim oWshShell, oExec, oClipboard 'oSys3
	Dim iRetryCount
	Dim sRetVal
	'REM changed, because Mercury.Clipboard is QTP specific and Clipboard object is not available in VBScript, made own utility
	'Set oSys3 = CreateObject("JSSys3.ops")
	'oSys3.SendTextCB("")
	'Set oClipboard = CreateObject("Mercury.Clipboard")
	'oClipboard.SetText ""
	Set oClipboard = CreateObject("TAUtility.TAClipboard")
	Set oWshShell = CreateObject("WScript.Shell")

	For iRetryCount = 1 to 5
		Set oExec = oWshShell.Exec(sKeePassFolder & "KPScript -c:clippassword " & """" & sDBFile & _
									""" -keyfile:""" & sKeyFile & """ -ref-Title:""" & sTitle & """")
		Do While oExec.Status = 0
			'WaitTime 0, 100
			WaitTime(100)
		Loop

		'sRetVal = ClipboardText()
		'oSys3.GetTextCB sRetVal
		'sRetVal = oClipboard.GetText()
		sRetVal = oClipboard.GetTextFromClipboard()
		If Not (sRetVal = "" Or IsNull(sRetVal) Or IsEmpty(sRetVal)) Then
			oClipboard.ClearClipboard()
			Exit For
		End If
	Next
	
	'Set oSys3 = Nothing
	Set oClipboard = Nothing
	Set oExec = Nothing
	Set oWshShell = Nothing

	KeePassPassword = sRetVal
End Function

'==============================================================================================
' Function/Sub: CheckKPKeyFile()
' Purpose: 
'==============================================================================================
Public Function CheckKPKeyFile(sKeyFileFolder, sUserNameOverride)
	Dim oFS, oFile, sKeyFile
	Dim bRetVal

	bRetval = False
	
	Set oFS = CreateObject("Scripting.FileSystemObject")
	If Trim(sUserNameOverride) = "" Or UCase(Trim(sUserNameOverride)) = "DEFAULT"  Then  
		sKeyFile = sKeyFileFolder & "KPKeyFile" & GetQCWindowsUser()
	Else
		sKeyFile = sKeyFileFolder & "KPKeyFile" & sUserNameOverride
	End If
	
	'Check that the keyfile is already there, and if not add it	
	if oFS.FileExists(sKeyFile) then
		bRetVal = True
	Elseif oFS.FileExists(sKeyFile & ".key") then
		bRetVal = True
	End If
	
	Set oFS = Nothing

	CheckKPKeyFile = bRetVal
End Function


'==============================================================================================
' Function/Sub:	Executes a command using cmd /c or cmd /k. Arguments:
' 				sCommand - command to be executed
' 				bWaitForFinish - if set to True then will invoke command and WaitTime otherwise will return immediately.
' 				bCloseOnFinish - if set to True then command window will close on finish otherwise will stay open. Useful for debug.
'==============================================================================================
Public Sub ExecuteCommand(sCommand, bWaitForFinish, bCloseOnFinish, sLogfile)
	Dim WshShell, oExec
	Dim sCmd

	Set WshShell = CreateObject("WScript.Shell")

	sCmd = "cmd /k "
	if bCloseOnFinish = True then
		sCmd = "cmd /c "
	end if

	if sLogfile = "" then
		Set oExec = WshShell.Exec(sCmd & " """ & sCommand & """")
	Else
		CheckForBackup sLogFile 'checking if the log file is not too big. If so, method will also perform backup.
		WshShell.Run "cmd /c echo . >> " & sLogFile, 1, True
		WshShell.Run "cmd /c echo %date% %time% START " & sCommand & " ***************** >> " & sLogFile, 1, True
		Set oExec = WshShell.Exec(sCmd & " """ & sCommand & " >> " & sLogFile & """")
	end if

	if bWaitForFinish = True then
		'WaitTime for the cmd to finish ...
		Do While oExec.Status = 0
			 'WaitTime 0,100
			 WaitTime(100)
		Loop

		if sLogfile <> "" then
			WshShell.Run "cmd /c echo %date% %time% FINISH " & sCommand & " ***************** >> " & sLogFile, 1, True
		end if
	end if

	Set oExec = nothing
	Set WshShell = nothing
End Sub

'==============================================================================================
' Function/Sub: 
' Purpose: 
'==============================================================================================
Public Function DictFromINI(sIniFileName)
	Dim dictSection, blnFoundSection, strSection
	Dim iEquals, sKey, sVal, i, sLine, oFileINI
	Dim oFS
	
	On Error Resume Next
	
	Set oFS = CreateObject("Scripting.Filesystemobject")
	
	blnFoundSection = False
	Err.Clear
	
	If oFS.FileExists(sIniFileName) Then
	
	    Set oFileINI = oFS.OpenTextFile(sIniFileName)
	    Set DictFromINI = CreateObject("Scripting.Dictionary")
	    
	    Do While Not oFileINI.AtEndOfStream
	        sLine = ""
	        sLine = Trim(oFileINI.ReadLine)
	        If sLine <> "" Then
	            If Left(sLine,1) <> ";" Then
	                If Left(sLine,1) = "[" Then
	                    blnFoundSection = True
	                    'Msgbox sLine & " section found"
	                    strSection = Left(sLine, Len(sLine) - 1)
	                    strSection = Right(strSection, Len(strSection) - 1)
	                    Set dictSection = CreateObject("Scripting.Dictionary")
	                    DictFromINI.Add UCase(strSection), dictSection
	                Else
	                    'key and value logic
	                    iEquals = InStr(1, sLine, "=")
	                    If (iEquals <= 1) Then
	                        'msgbox "error: the following line is invalid " & sLine
	                    Else
	                        'we've found a valid line
	                        sKey = Left(sLine, iEquals - 1)
	                        sVal = Right(sLine, Len(sLine) - iEquals)
	                        	                        
	                        Err.Clear
	                        dictSection.Add Trim(LCase(sKey)), Trim(sVal)
	                        If Err.Number <> 0 Then
	                            'msgbox "unable to add to dictionary object"
	                        End If
	                        'msgbox strSection & " " & sKey & ";;;;" & sVal
	                        
	                    'key and value logic end if
	                    End If
	                End If
	            End If
	        End If
	    Loop
	    
	    oFileINI.Close
	    Set oFileINI = Nothing
	    
	    If blnFoundSection = False Then
	        Set DictFromINI = CreateObject("Scripting.Dictionary")
	    End If
	
	Else
	    Set DictFromINI = CreateObject("Scripting.Dictionary")
	End If
	
	Set oFS = Nothing
End Function

'==============================================================================================
' Function/Sub: 
' Purpose: 
'==============================================================================================
Public Function DictFromXLS(sFile)
	Dim oFS, oExcel, oWorkbook, oSheet
	Dim dictSection
	Dim iRow, iPosition, bValid
	Dim sCellText, sValidFrom, sValidTo
	Dim dateFrom, dateTo
	
	Set DictFromXLS = CreateObject("Scripting.Dictionary")
	Set oFS = CreateObject("Scripting.Filesystemobject")
	Set oExcel = CreateObject("Excel.Application")
	oExcel.Visible = False
	
	If oFS.FileExists(sFile) Then
		Set oWorkbook = oExcel.Workbooks.Open(sFile)
		
		For Each oSheet In oWorkbook.Sheets
			Set dictSection = CreateObject("Scripting.Dictionary")
			
			For iRow = 1 To oSheet.Cells(oSheet.Rows.Count, 1).End(-4162).Row				
				sCellText = oSheet.Cells(iRow, 1).Text				
				bValid = True
				' for MAPPING sheet take only current valid records
				If oSheet.Name = "MAPPING" Then
					sValidFrom = oSheet.Cells(iRow, 3).Text
					sValidTo = oSheet.Cells(iRow, 4).Text
					If IsDate(sValidFrom) Then
						dateFrom = CDate(sValidFrom)
					Else
						dateFrom = CDate("01/01/1900")
					End If
					If IsDate(sValidTo) Then
						dateTo = CDate(sValidTo)
					Else
						dateTo = CDate("31/12/2099")
					End If					
					If Not(Date() >= dateFrom And Date <= dateTo) Then
						bValid = False
					End If
				End If
				' add records to dict; if 2 with the same name, then last one will overwrite previous one
				If bValid Then				
					If sCellText = "" Then
						'do nothing
					ElseIf InStr(sCellText, "=") > 0 Then
						iPosition = InStr(sCellText, "=")
						'dictSection.Add Trim(LCase(Left(sCellText, iPosition - 1))), Trim(Mid(sCellText, iPosition + 1))
						dictSection.Item(Trim(LCase(Left(sCellText, iPosition - 1)))) = Trim(Mid(sCellText, iPosition + 1))
					Else
						'dictSection.Add Trim(LCase(sCellText)), Trim(oSheet.Cells(iRow, 2).Text)
						dictSection.Item(Trim(LCase(sCellText))) = Trim(oSheet.Cells(iRow, 2).Text)
					End If
				End If
			Next

			'DictFromXLS.Add UCase(oSheet.Name), dictSection
			Set DictFromXLS.Item(UCase(oSheet.Name)) = dictSection
			Set dictSection = Nothing
		Next
	End If
	
	oWorkbook.Close
	oExcel.Quit
	
	Set oExcel = Nothing
	Set oFS = Nothing
End Function

'==============================================================================================
' Function/Sub: 
' Purpose: 
'==============================================================================================
Public Function GetScreenResolution()
	Dim oIE, iWidth, iHeight
	Set oIE = CreateObject("InternetExplorer.Application")
	oIE.Navigate("about:blank")
	Do Until oIE.readyState = 4
		'WaitTime(1)
		WaitTime(100)
	Loop
	iWidth = oIE.document.ParentWindow.screen.width
	iHeight = oIE.document.ParentWindow.screen.height
	oIE.Quit
	Set oIE = Nothing
	GetScreenResolution = Array(iWidth,iHeight)
End Function

'==============================================================================================
' Function/Sub: PdfCompare() - return true if files match
' Purpose:		Compare pdf files.  First convert them to text files using XPDF's pdftotext.exe
'				Then conpare the text files allowing for some text to be ignored.
'				sIgnoreText is in the form of row,column,length|row,column,length|row,column,length etc
'==============================================================================================
Public Function PdfCompare(sXPDFFolder, sPDFFile1, sPDFFile2, sIgnoreText, sPatternList, sLog)
	Dim bRetVal
	Dim sTextFile1, sTextFile2
	Dim oFS

	bRetVal = False

	Set oFS = CreateObject("Scripting.FileSystemObject")

	'Create the textfile names
'	sTextFile1 = oFS.GetParentFolderName(sPDFFile1) & "\" & oFS.GetBaseName(sPDFFile1) & ".txt"
'	sTextFile2 = oFS.GetParentFolderName(sPDFFile2) & "\" & oFS.GetBaseName(sPDFFile2) & ".txt"

	sTextFile1 = PATH_RESOURCES & oFS.GetBaseName(sPDFFile1) & ".txt"
	sTextFile2 = PATH_RESOURCES & oFS.GetBaseName(sPDFFile2) & ".txt"

	'Convert PDFs to Text
	PdfToText sXPDFFolder, sPDFFile1, sTextFile1
	PdfToText sXPDFFolder, sPDFFile2, sTextFile2

	'Compare the text files
	bRetVal = TextFileCompare(sTextFile1, sTextFile2, sIgnoreText, sPatternList, sLog)

	PdfCompare = bRetVal
End Function

'==============================================================================================
' Function/Sub: PdfToText()
' Purpose: 		Use the XPDF program to convert a pdf file to a text file. 
'==============================================================================================
Public Sub PdfToText(sXPDFFolder, sPDFFile, sOutputFile)
	Dim oWshShell, oExec

	Set oWshShell = CreateObject("WScript.Shell")
	Set oExec = oWshShell.Exec(sXPDFFolder & "\pdftotext.exe -layout -nopgbrk """ & sPDFFile & """ """ & sOutputFile & """")

	Do While oExec.Status = 0
		'WaitTime 0, 100
		WaitTime(100)
	Loop

	Set oExec = Nothing
	Set oWshShell = Nothing
End Sub


'==============================================================================================
' Function/Sub:  Returns the Full Path of the file if it is found
'==============================================================================================
Function FindFile(sFilename, sStartFolder)
	Dim sRetval
	Dim oFS, oFile, oFolder, oSubFolder

	sRetval = ""
	Set oFS = CreateObject("Scripting.FileSystemObject")

	'*PREREQ
	If Not oFS.FolderExists(sStartFolder) Then
		Exit Function
	End If
	'*
	
	Set oFolder = oFS.GetFolder(sStartFolder)
	For Each oFile in oFolder.Files
		If UCase(oFile.Name) = UCase(sFilename) Then
			sRetval = oFile.Path	' & "\" & sFilename
			Exit For
		End If
	Next

	If sRetval = "" Then
		For Each oSubFolder in oFolder.SubFolders 
			sRetval = FindFile(sFilename, oSubFolder.Path) ' & "\" & oSubFolder.Name)
			if sRetval <> "" then Exit For
		Next
	End If
			
	Set oFolder = Nothing
	Set oFS = Nothing

	FindFile = sRetval
End Function

'==============================================================================================
' Function/Sub:  
'==============================================================================================
Public Function FileFromSpecification(sFileSpecification)
	Dim sRetval
	Dim dictFile', oQC
	Dim iCount

	sRetval = ""
	Set dictFile = CreateObject("Scripting.Dictionary")
	'Set oQC = New clsQC

	'Check if the file is attached, or in the filesystem
	if SimpleDictKeyValuePairs(dictFile, sFileSpecification, ",", "=") > 0 Then 
		If dictFile.Exists("Type") Then
			Select Case UCase(dictFile("Type"))
'			Case "ATTACHMENTBYPATTERN"
'				If oQC.IsQCRun() = True Then 
'					sRetval = oQC.GetAttachmentFileFromQCByPattern(dictFile("Pattern"), PATH_RESOURCES)
'				End If
			Case "ATTACHMENT"
				If oQC.IsQCRun() = True Then 
					'sRetval = oQC.GetAttachmentFileFromQC(dictFile("Prefix"), dictFile("Extension"), PATH_RESOURCES)
					If dictFile.Exists("Prefix") And dictFile.Exists("Extension") Then 
						sRetval = oQC.GetAttachmentFileFromQCByPattern(dictFile("Prefix") & ".*" & dictFile("Extension"), PATH_RESOURCES)
					ElseIf dictFile.Exists("Pattern") then
						sRetval = oQC.GetAttachmentFileFromQCByPattern(dictFile("Pattern"), PATH_RESOURCES)
					End If
				End If
			Case "FILESYSTEM"
				sRetval = dictFile("Filename")
			Case "RESOURCE"
				If oQC.IsQCRun() = True Then 
					sRetval = oQC.GetResourceFileFromQC(dictFile("Name"))
				End If
			Case Else
				
			End Select
		End If
	End if

'	Set oQC = Nothing
	Set dictFile = Nothing
	FileFromSpecification = sRetval
End Function

'--------------------------------------------------------------------------------------------------
' Function/Sub: GetShortString
' Purpose: 
'--------------------------------------------------------------------------------------------------
Public Function GetShortString(sLongString, iLength)
	Dim sRetVal
	Dim sNewString, sPartString, iLenPart, sTempString
	Dim i, j, iParts

	sTempString = sLongString & string(iLength-1, chr(0))
	iParts = int(Len(sTempString)/iLength)
	sTempString = left(sTempString, iParts * iLength)
	sPartString = left(sTempString, iLength)

	For i = 1 to iParts-1
		sNewString = ""
		For j = 1 to iLength
			sNewString = sNewString & chr((Asc(mid(sPartString, j, 1)) + Asc(mid(sTempString, i*iLength + j, 1))) mod 255) 
		Next
		sPartString = sNewString
	Next

	GetShortString = sNewString
End Function

'--------------------------------------------------------------------------------------------------
' Function/Sub: Base64Encode
' Purpose: 
'--------------------------------------------------------------------------------------------------
Function Base64Encode(inData)
  'rfc1521
  '2001 Antonin Foller, Motobit Software, http://Motobit.cz
  'Const Base64 = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789+/"
	Const Base64 = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789"
  Dim cOut, sOut, I
  
  'For each group of 3 bytes
  For I = 1 To Len(inData) Step 3
    Dim nGroup, pOut, sGroup
    
    'Create one long from this 3 bytes.
    nGroup = &H10000 * Asc(Mid(inData, I, 1)) + &H100 * MyASC(Mid(inData, I + 1, 1)) + MyASC(Mid(inData, I + 2, 1))
    
    'Oct splits the long To 8 groups with 3 bits
    nGroup = Oct(nGroup)
    
    'Add leading zeros
    nGroup = String(8 - Len(nGroup), "0") & nGroup
    
    'Convert To base64
    pOut = Mid(Base64, CLng("&o" & Mid(nGroup, 1, 2)) + 1, 1) + _
      Mid(Base64, CLng("&o" & Mid(nGroup, 3, 2)) + 1, 1) + _
      Mid(Base64, CLng("&o" & Mid(nGroup, 5, 2)) + 1, 1) + _
      Mid(Base64, CLng("&o" & Mid(nGroup, 7, 2)) + 1, 1)
    
    'Add the part To OutPut string
    sOut = sOut + pOut
    
    'Add a new line For Each 76 chars In dest (76*3/4 = 57)
    'If (I + 2) Mod 57 = 0 Then sOut = sOut + vbCrLf
  Next
'  Select Case Len(inData) Mod 3
'    Case 1: '8 bit final
'      sOut = Left(sOut, Len(sOut) - 2) + "=="
'    Case 2: '16 bit final
'      sOut = Left(sOut, Len(sOut) - 1) + "="
'  End Select
  Base64Encode = sOut
End Function

'==============================================================================================
'==============================================================================================
Function MyASC(OneChar)
  If OneChar = "" Then MyASC = 0 Else MyASC = Asc(OneChar)
End Function

'==============================================================================================
' Function/Sub: SaveTextToFile
' Purpose: 
'==============================================================================================
Public Function SaveTextToFile(sText, sFilename)
	Dim oFS, oFile
	
	Set oFS = CreateObject("Scripting.FileSystemObject")
	Set oFile = oFS.OpenTextFile(sFilename, 2, True)
	oFile.Write sText
	oFile.Close
	
	Set oFile = Nothing
	Set oFS = Nothing

End Function

'==============================================================================================
' Function/Sub: LoadObjectRepository()
' Purpose: 
'==============================================================================================
'REM QTP Specific function
Public Function LoadObjectRepository(sPath) 
	Dim oQTP, oQTPRepositories, sActionName
	Dim oFS, bRetval

	bRetVal = False
	
	Set oFS = CreateObject("Scripting.FileSystemObject")
	If  oFS.FileExists(sPath)= True Then				'Only run if file exists.
		bRetVal = True
		Set oQTP = CreateObject("QuickTest.Application")
		sActionName = Environment("ActionName") 'getting the action name
		Set oQTPRepositories = oQTP.Test.Actions(sActionName).ObjectRepositories

		If oQTPRepositories.Find(sPath) = -1 Then ' If the repository cannot be found in the collection
			oQTPRepositories.Add sPath ' Add the repository to the collection
		End If
		
		Set oQTPRepositories = Nothing
		Set oQTP = Nothing
	End If

	Set oFS = Nothing
	LoadObjectRepository = bRetVal
	
End Function

'==============================================================================================
' Function/Sub: Touch()
' Purpose: 
'==============================================================================================
Public Sub Touch(sFolderPath, sFileName) 
	Dim oApp, oFolder, oFile

	Set oApp = CreateObject("Shell.Application") 
	Set oFolder = oApp.NameSpace(sFolderPath) 
	Set oFile = oFolder.ParseName(sFileName) 
	 
	oFile.ModifyDate = CStr(now) 
	
	set oFile = nothing 
	set oFolder = nothing 
	set oApp = nothing 
End Sub 

'==============================================================================================
' Function/Sub: TextFileTextReplace(sInFile, sOutFile, sFindText, sReplaceText, sOutput) - return false on error
' Purpose: 		Returns True if the output file is written, False if the input file does not exist or the output cannot be created 	
'==============================================================================================
Public Function TextFileTextReplace(sInFile, sOutFile, sFindText, sReplaceText, sLog)
	Dim bRetVal, iExistRetry, iReadRetry
	Dim oFS, oInFile, oOutFile
	Dim sInText, sOutText

	bRetVal = True
	sInText = Empty
	iExistRetry = 1
	iReadRetry = 1

    Set oFS = CreateObject("Scripting.FileSystemObject")

	Do While Not oFS.FileExists(sInFile) and iExistRetry < 4
		sLog = sLog & "Warning: Input file " & sInFile &  " not found at attempt " & iRetry & vbLf
		'WaitTime 1
		WaitTime(100)
		iExistRetry =  iExistRetry + 1
	Loop

	If oFS.FileExists(sInFile) Then
		'Put in error handling to make sure we get the data in the file
		On error resume next 
		Do while isempty(sInText) and iReadRetry < 6
			Set oInFile = oFS.OpenTextFile(sInFile, 1, False)	'Read
			sInText = oInFile.ReadAll
			If err.number <> 0  Then
				sLog = sLog & "Warning: Failed at attempt " & iRetry & " to read file " & sInFile & vbLf
				'WaitTime 1
				WaitTime(100)
			End If
			oInFile.Close
			Set oInFile = Nothing
			iReadRetry = iReadRetry + 1
		Loop
		On error goto 0 

		If  not IsEmpty(sInText) Then
			sOutText = Replace(sInText, sFindText, sReplaceText)
			sLog = sLog & "Debug: Input file " & sInFile & " read." & vbLf
			
			If oFS.FolderExists(oFS.GetParentFolderName(sOutFile)) Then
				Set oOutFile = oFS.OpenTextFile(sOutFile, 2, True)	'Write
				oOutFile.write sOutText
				oOutFile.Close
				sLog = sLog & "Debug: Find " & sFindText & " and replace with " & sReplaceText & vbLf
				sLog = sLog & "Debug: Output  file " & sOutFile & " created." & vbLf
				Set oOutFile = Nothing
			Else
				sLog = sLog & "Error: Output file path " & oFS.GetParentFolderName(sOutFile) & " not found." & vbLf
				bRetVal = False
			End If
		Else
			sLog = sLog & "Error: Failed reading input file " & sInFile & vbLf
			bRetVal = False
		End If
	Else
		sLog = sLog & "Error: Input file " & sInFile & " not found." & vbLf
		bRetVal = False
	End If 

    Set oFS = Nothing
	TextFileTextReplace = bRetVal

End Function

'-------------------------------------------------------------------------------------------------------------------------------------
' Function/Sub: BuildIgnoreTextListFromFile(sFilename, sPatternList, sDelimiter, sOutput) 
' Purpose: 		return sIgnoreText in the form of row,column,length|row,column,length|row,column,length etc
'-------------------------------------------------------------------------------------------------------------------------------------
Function BuildIgnoreTextListFromFile(sFilename, sPatternList, sDelimiter, sLog)
	Dim i, iLineNumber
	Dim sLine
	Dim sRet', sOutput
	Dim sType, sPattern, iMaxLength
	Dim arrPatternList, arrSpec
	Dim arrType(), arrPattern(), arrMaxLength()
	Dim oFS, oFile

	sRet = ""
	'sOutput = ""

	'Build the arrays of patterns 
	arrPatternList = split(sPatternList, "|")
	ReDim arrType(ubound(arrPatternList))
	ReDim arrPattern(ubound(arrPatternList))
	ReDim arrMaxLength(ubound(arrPatternList))

	For i = 0 to ubound(arrPatternList)
		arrSpec = split(arrPatternList(i), ",")
	
		If UBound(arrSpec) = 2 Then
			arrType(i) = arrSpec(0)
			arrPattern(i) = arrSpec(1)
			arrMaxLength(i) = Cint("0" & arrSpec(2))
		
			Select Case UCase(arrType(i))
			Case "TODAY"
				arrPattern(i) = Replace(arrPattern(i), "YYYY", Year(Now))
				arrPattern(i) = Replace(arrPattern(i), "DD", Right("0" & Day(Now), 2))
				arrPattern(i) = Replace(arrPattern(i), "D", CStr(Day(Now)))
				arrPattern(i) = Replace(arrPattern(i), "MONTH", MonthName(Month(Now), False))
				arrPattern(i) = Replace(arrPattern(i), "MMM", MonthName(Month(Now), True))
				arrPattern(i) = Replace(arrPattern(i), "MM", Right("0" & Month(Now), 2))
			Case "REGEX"
				'do Nothing to the pattern
			Case Else
				sLog = sLog &  "Error: Unrecognised pattern type: " & arrType(i) & " in pattern: " & arrPatternList(i) & vbnewline
				'msgbox sLog
				BuildIgnoreTextListFromFile = "" : Exit Function
			End Select
		Else
			sLog = sLog &  "Error: Invalid number (" & UBound(arrSpec)+1 & ") of specs in pattern: " & arrPatternList(i) & vbnewline
			'msgbox sLog
			BuildIgnoreTextListFromFile = "" : Exit Function
		End If
	Next

	'Now build the IgnoreText for all lines in the file
    Set oFs = CreateObject("Scripting.FileSystemObject")
    Set oFile = oFs.OpenTextFile(sFilename, 1, False)

	iLineNumber = 1
	Do While Not oFile.AtEndOfStream
		sLine = oFile.ReadLine
		For i = 0 to Ubound(arrPattern)
			sRet = sRet & BuildIgnoreTextListFromString(sLine, arrPattern(i), arrMaxLength(i), "|" & iLineNumber & ",", sLog)
		Next

		iLIneNumber = iLineNumber + 1
	Loop

    Set oFile = Nothing
    Set oFs = Nothing

	If len(sRet) > 0 Then
		sRet = mid(sRet, 2)
	End If
	BuildIgnoreTextListFromFile = sRet

End Function

'-------------------------------------------------------------------------------------------------------------------------------------
' Function/Sub: BuildIgnoreTextListFromFile2(sFilename, sPatternList, sDelimiter, sOutput) 
' Purpose: 		return sIgnoreText in the form of row,column,length|row,column,length|row,column,length etc
'-------------------------------------------------------------------------------------------------------------------------------------
Function BuildIgnoreTextListFromFile2(sFilename, sPatternList, sDelimiter, arrNewToOld, sLog)
	Dim i, iLineNumber
	Dim sLine
	Dim sRet', sOutput
	Dim sType, sPattern, iMaxLength
	Dim arrPatternList, arrSpec
	Dim arrType(), arrPattern(), arrMaxLength()
	Dim oFS, oFile

	sRet = ""
	'sOutput = ""

	'Build the arrays of patterns 
	arrPatternList = split(sPatternList, "|")
	ReDim arrType(ubound(arrPatternList))
	ReDim arrPattern(ubound(arrPatternList))
	ReDim arrMaxLength(ubound(arrPatternList))

	For i = 0 to ubound(arrPatternList)
		arrSpec = split(arrPatternList(i), ",")
	
		If UBound(arrSpec) = 2 Then
			arrType(i) = arrSpec(0)
			arrPattern(i) = arrSpec(1)
			arrMaxLength(i) = Cint("0" & arrSpec(2))
		
			Select Case UCase(arrType(i))
			Case "TODAY"
				arrPattern(i) = Replace(arrPattern(i), "YYYY", Year(Now))
				arrPattern(i) = Replace(arrPattern(i), "DD", Right("0" & Day(Now), 2))
				arrPattern(i) = Replace(arrPattern(i), "D", CStr(Day(Now)))
				arrPattern(i) = Replace(arrPattern(i), "MONTH", MonthName(Month(Now), False))
				arrPattern(i) = Replace(arrPattern(i), "MMM", MonthName(Month(Now), True))
				arrPattern(i) = Replace(arrPattern(i), "MM", Right("0" & Month(Now), 2))
			Case "REGEX"
				'do Nothing to the pattern
			Case Else
				sLog = sLog &  "Error: Unrecognised pattern type: " & arrType(i) & " in pattern: " & arrPatternList(i) & vbnewline
				'msgbox sLog
				BuildIgnoreTextListFromFile2 = "" : Exit Function
			End Select
		Else
			sLog = sLog &  "Error: Invalid number (" & UBound(arrSpec)+1 & ") of specs in pattern: " & arrPatternList(i) & vbnewline
			'msgbox sLog
			BuildIgnoreTextListFromFile2 = "" : Exit Function
		End If
	Next

	'Now build the IgnoreText for all lines in the file
    Set oFs = CreateObject("Scripting.FileSystemObject")
    Set oFile = oFs.OpenTextFile(sFilename, 1, False)

	iLineNumber = 1
	Do While Not oFile.AtEndOfStream
		sLine = oFile.ReadLine
		For i = 0 to Ubound(arrPattern)
			sRet = sRet & BuildIgnoreTextListFromString2(sLine, arrPattern(i), arrMaxLength(i), "|" & iLineNumber & ",", arrNewToOld, sLog)
			'sRet = sRet & BuildIgnoreTextListFromString(sLine, arrPattern(i), arrMaxLength(i), "|" & arrNewToOld(iLineNumber) & "(" & iLineNumber & ")"  & ",", sLog)
		Next

		iLIneNumber = iLineNumber + 1
	Loop

    Set oFile = Nothing
    Set oFs = Nothing

	If len(sRet) > 0 Then
		sRet = mid(sRet, 2)
	End If
	BuildIgnoreTextListFromFile2 = sRet

End Function

'-------------------------------------------------------------------------------------------------------------------------------------
' Function/Sub: BuildIgnoreTextListFromArray(arrLineText, sPatternList, sDelimiter, sOutput) 
' Purpose: 		return sIgnoreText in the form of row,column,length|row,column,length|row,column,length etc
'-------------------------------------------------------------------------------------------------------------------------------------
Function BuildIgnoreTextListFromArray(arrLineText, sPatternList, sDelimiter, sLog)
	Dim i, iLineNumber
	Dim sLine
	Dim sRet', sOutput
	Dim sType, sPattern, iMaxLength
	Dim arrPatternList, arrSpec
	Dim arrType(), arrPattern(), arrMaxLength()
	Dim oFS, oFile

	sRet = ""
	'sOutput = ""

	'Build the arrays of patterns 
	arrPatternList = split(sPatternList, "|")
	ReDim arrType(ubound(arrPatternList))
	ReDim arrPattern(ubound(arrPatternList))
	ReDim arrMaxLength(ubound(arrPatternList))

	For i = 0 to ubound(arrPatternList)
		arrSpec = split(arrPatternList(i), ",")
	
		If UBound(arrSpec) = 2 Then
			arrType(i) = arrSpec(0)
			arrPattern(i) = arrSpec(1)
			arrMaxLength(i) = Cint("0" & arrSpec(2))
		
			Select Case UCase(arrType(i))
			Case "TODAY"
				arrPattern(i) = Replace(arrPattern(i), "YYYY", Year(Now))
				arrPattern(i) = Replace(arrPattern(i), "DD", Right("0" & Day(Now), 2))
				arrPattern(i) = Replace(arrPattern(i), "D", CStr(Day(Now)))
				arrPattern(i) = Replace(arrPattern(i), "MONTH", MonthName(Month(Now), False))
				arrPattern(i) = Replace(arrPattern(i), "MMM", MonthName(Month(Now), True))
				arrPattern(i) = Replace(arrPattern(i), "MM", Right("0" & Month(Now), 2))
			Case "REGEX"
				'do Nothing to the pattern
			Case Else
				sLog = sLog &  "Error: Unrecognised pattern type: " & arrType(i) & " in pattern: " & arrPatternList(i) & vbnewline
				'msgbox sLog
				BuildIgnoreTextListFromArray = "" : Exit Function
			End Select
		Else
			sLog = sLog &  "Error: Invalid number (" & UBound(arrSpec)+1 & ") of specs in pattern: " & arrPatternList(i) & vbnewline
			'msgbox sLog
			BuildIgnoreTextListFromArray = "" : Exit Function
		End If
	Next

	'Now build the IgnoreText for all lines in the array
	For iLineNumber = 0 To UBound(arrLineText)
		sLine = arrLineText(iLineNumber)
		For i = 0 to Ubound(arrPattern)
			sRet = sRet & BuildIgnoreTextListFromString(sLine, arrPattern(i), arrMaxLength(i), "|" & iLineNumber & ",", sLog)
		Next
	next

	If len(sRet) > 0 Then
		sRet = mid(sRet, 2)
	End If
	BuildIgnoreTextListFromArray = sRet

End Function

'-------------------------------------------------------------------------------------------------------------------------------------
' Function/Sub: BuildIgnoreTextListFromString(sString, sPattern, iMaxLength, sPrefix, sOutput)
' Purpose: 		return sIgnoreText in the form of row,column,length|row,column,length|row,column,length etc
'-------------------------------------------------------------------------------------------------------------------------------------
Function BuildIgnoreTextListFromString(sString, sPattern, iMaxLength, sPrefix, sLog)
	Dim sRet', sOutput
	Dim oRegExp, oMatch, oMatches
	Dim iPosition

	sRet = ""

	Set oRegExp = New RegExp		' Create a regular expression.
	oRegExp.Pattern = sPattern      ' Set pattern.
	oRegExp.IgnoreCase = True       ' Set case insensitivity.
	oRegExp.Global = True         	' Set global applicability.
	
	Set oMatches = oRegExp.Execute(sString)   ' Execute search.
	For Each oMatch in oMatches      ' Iterate Matches collection.
		iPosition = oMatch.FirstIndex + 1
		sLog = sLog & "Pattern matched at line " & Mid(sPrefix,2) & " position " & iPosition & ". Pattern value is '" & oMatch.Value & "'." & vbLf
		sRet = sRet & sPrefix & iPosition & "," 
		If (Len(sString) > iPosition + iMaxLength - 1) And (iMaxLength > 0)  Then 	'Check if the end of the line is reached
			sRet = sRet & iMaxLength
		End If
	Next

	Set oMatches = Nothing
	Set oRegExp = Nothing
	
	'if sLog <> "" then msgbox sLog
	BuildIgnoreTextListFromString = sRet

End Function

'-------------------------------------------------------------------------------------------------------------------------------------
' Function/Sub: BuildIgnoreTextListFromString2(sString, sPattern, iMaxLength, sPrefix, sOutput)
' Purpose: 		return sIgnoreText in the form of row,column,length|row,column,length|row,column,length etc
'-------------------------------------------------------------------------------------------------------------------------------------
Function BuildIgnoreTextListFromString2(sString, sPattern, iMaxLength, sPrefix, arrNewToOld, sLog)
	Dim sRet', sOutput
	Dim oRegExp, oMatch, oMatches
	Dim iPosition

	sRet = ""

	Set oRegExp = New RegExp		' Create a regular expression.
	oRegExp.Pattern = sPattern      ' Set pattern.
	oRegExp.IgnoreCase = True       ' Set case insensitivity.
	oRegExp.Global = True         	' Set global applicability.
	
	Set oMatches = oRegExp.Execute(sString)   ' Execute search.
	For Each oMatch in oMatches      ' Iterate Matches collection.
		iPosition = oMatch.FirstIndex + 1
		sLog = sLog & "Pattern matched at line " & arrNewToOld(CInt(Mid(sPrefix,2))) & "(" & CInt(Mid(sPrefix,2)) & ")" & " position " & iPosition & ". Pattern value is '" & oMatch.Value & "'." & vbLf
		sRet = sRet & sPrefix & iPosition & "," 
		If (Len(sString) > iPosition + iMaxLength - 1) And (iMaxLength > 0)  Then 	'Check if the end of the line is reached
			sRet = sRet & iMaxLength
		End If
	Next

	Set oMatches = Nothing
	Set oRegExp = Nothing
	
	'if sLog <> "" then msgbox sLog
	BuildIgnoreTextListFromString2 = sRet

End Function

'-------------------------------------------------------------------------------------------------------------------------------------
' Function/Sub: PatternFirstMatch(sString, sPattern, sOutput)
' Purpose: 		look for a pattern in the string and return the first match
'-------------------------------------------------------------------------------------------------------------------------------------
Function PatternFirstMatch(sString, sPatternSpec, sLog)
	Dim sRet
	Dim oRegExp, oMatch, oMatches
	Dim arrSpec, sType, sPattern 

	sRet = ""
	arrSpec = split(sPatternSpec, ",")
	If UBound(arrSpec) = 1 Then
		sType = arrSpec(0)
		sPattern = arrSpec(1)
	
		Select Case UCase(sType)
		Case "TODAY"
			sPattern = Replace(sPattern, "YYYY", Year(Now))
			sPattern = Replace(sPattern, "DD", Right("0" & Day(Now), 2))
			sPattern = Replace(sPattern, "D", CStr(Day(Now)))
			sPattern = Replace(sPattern, "MONTH", MonthName(Month(Now), False))
			sPattern = Replace(sPattern, "MMM", MonthName(Month(Now), True))
			sPattern = Replace(sPattern, "MM", Right("0" & Month(Now), 2))
		Case "REGEX"
			'do Nothing to the pattern
		Case Else
			sLog = sLog &  "Error: Unrecognised pattern type: " & sType & " in pattern: " & sPattern & vbnewline
			'msgbox sLog
			BuildIgnoreTextListFromFile = "" : Exit Function
		End Select
	Else
		sLog = sLog &  "Error: Invalid number (" & UBound(arrSpec)+1 & ") of specs in pattern: " & sPatternSpec & vbnewline
		'msgbox sLog
		PatternFirstMatch = "" : Exit Function
	End If

'stop
	Set oRegExp = New RegExp		' Create a regular expression.
	oRegExp.Pattern = sPattern      ' Set pattern.
	oRegExp.IgnoreCase = True       ' Set case insensitivity.
	oRegExp.Global = True         	' Set global applicability.

	Set oMatches = oRegExp.Execute(sString)   ' Execute search.
	For Each oMatch in oMatches      ' Iterate Matches collection.
		sLog = sLog & "Pattern '" & sPattern & "' matched. Pattern value is '" & oMatch.Value & "'." & vbLf
		sRet = oMatch.Value
		Exit For
	Next

	Set oMatch = Nothing
	Set oMatches = Nothing
	Set oRegExp = Nothing
	
	PatternFirstMatch = sRet

End Function

'==============================================================================================
' Function/Sub: TextFileCompare(sTextFile1, sTextFile2, sIgnoreText) - return true if files match
' Purpose: 		sIgnoreText is in the form of row,column,length|row,column,length|row,column,length etc
'				sPatternList is in the form of Type,Pattern,MaxLength|Type,Pattern,MaxLength etc
'				where Type is TODAY and Pattern example is DD-MMM-YYYY (regular expressions can still be embedded)
'				and   Type is REGEX and Pattern is any regular expression
'==============================================================================================
Public Function TextFileCompare(sTextFile1, sTextFile2, sIgnoreText, sPatternList, sLog)
	Dim bRetVal
	Dim oFS, oTextFile1, oTextFile2
	Dim sLine1, sLine2, iLine, sMasked1, sMasked2
	Dim arrIgnoreText, oDictIgnoreTextColumn, oDictIgnoreTextLength
	Dim sEntry, arrEntry, iOccurrence, iLineOcc, sIgnoreText2
	Dim iCount1, iCount2 

	bRetVal = True

    'First check the line count of the files.
	iCount1 = CountFileLines(sTextFile1)
	iCount2 = CountFileLines(sTextFile2)
	If iCount1<> iCount2 Then
		sLog = sLog & "E: Line counts are different for compared files. " & iCount1 & " vs " & iCount2 & vbLf
		bRetVal = False
	Else
		sLog = sLog & "I: Line counts are the same for compared files. " & iCount1 & " vs " & iCount2 & vbLf
	End if

	' This is to simulate line count differences for testing ...
'	sLog = sLog & "I: Simulating different line counts. " & vbLf
'	Call AddBlankLinesToFile(sTextFile2, 999, 1)
'	iCount1 = CountFileLines(sTextFile1)
'	iCount2 = CountFileLines(sTextFile2)
'	If iCount1<> iCount2 Then
'		sLog = sLog & "E: Line counts are different for compared files. " & iCount1 & " vs " & iCount2 & vbLf
'		bRetVal = False
'	Else
'		sLog = sLog & "I: Line counts are the same for compared files. " & iCount1 & " vs " & iCount2 & vbLf
'	End if

    Set oFs = CreateObject("Scripting.FileSystemObject")
    Set oTextFile1 = oFs.OpenTextFile(sTextFile1, 1, False)
    Set oTextFile2 = oFs.OpenTextFile(sTextFile2, 1, False)
	
    Set oDictIgnoreTextColumn = CreateObject("Scripting.Dictionary")
    Set oDictIgnoreTextLength = CreateObject("Scripting.Dictionary")

	'Derive the ignore text list associated with the patternlist and add it to the ignoretext.
	'Use sTextFile2 because by convention that is the newly generated file with Today's dates
	sIgnoreText = sIgnoreText & "|"& BuildIgnoreTextListFromFile(sTextFile2, sPatternList, "|", sLog)
	If sIgnoreText <> "|" Then	'Have some ignoring to do ...
		If mid(sIgnoreText,1,1) = "|" Then
			sIgnoreText = mid(sIgnoreText,2)
		End If

		If mid(sIgnoreText,len(sIgnoreText),1) = "|" Then
			sIgnoreText = mid(sIgnoreText,1,len(sIgnoreText)-1)
		End If

		arrIgnoreText = split(sIgnoreText, "|")
		For each sEntry in arrIgnoreText
			arrEntry = split(sEntry, ",")
			'Have to allow for multiple ignored text on a line - add 1000 to the line number each subsequent time it appears
			'Eg if the same line, say 2, has 5 bits of text to ignore then they will stored as 2, 1002, 2002, 3002 and 4002
			For iOccurrence = 0 to 10
				iLineOcc = arrEntry(0) + iOccurrence*1000
				If  Not oDictIgnoreTextColumn.Exists(iLineOcc) Then Exit For
			Next
			
			'Check if the start and length to ignore are specified.
			If UBound(arrEntry) > 0 Then
				oDictIgnoreTextColumn.Add CInt(iLineOcc), CInt("0" & arrEntry(1))
				If UBound(arrEntry) > 1 Then
					oDictIgnoreTextLength.Add CInt(iLineOcc), CInt("0" & arrEntry(2))
				Else
					oDictIgnoreTextLength.Add CInt(iLineOcc), 0
				End If
			Else
				oDictIgnoreTextColumn.Add CInt(iLineOcc), 0
				oDictIgnoreTextLength.Add CInt(iLineOcc), 0
			End If

		Next
	End If
	
	iLine = 1
	Do While (iLine <= iCount1) And (iLine <= iCount2) 'Not oTextFile1.AtEndOfStream
		sLine1 = oTextFile1.ReadLine
		sLine2 = oTextFile2.ReadLine
		'Allow for the odd case where there is a comma without a following space.  Add a space after all commas. 
		sLine1 = Replace(sLine1, ",", ", ")
		sLine2 = Replace(sLine2, ",", ", ")
		'Mask the ignored text - equalise spaces and add @ to the start and end of each line to avoid checking for boundary conditions
		sMasked1 = "@" & EqualiseSpaces(sLine1, sLine2, iLine, 4, sLog) & "@"
		sMasked2 = "@" & sLine2 & "@"
		For iOccurrence = 0 to 10
			iLineOcc = iLine + iOccurrence*1000
			If oDictIgnoreTextColumn.Exists(iLineOcc) Then
				'Check the line is long enough
					If len(sLine1) >= oDictIgnoreTextColumn(iLine)+oDictIgnoreTextLength(iLineOcc)-1  Then
						If oDictIgnoreTextLength(iLineOcc) = 0 Then
							'leave only the line content up to the column position
							sMasked1 = Left(sMasked1, oDictIgnoreTextColumn(iLineOcc))
							sMasked2 = Left(sMasked2, oDictIgnoreTextColumn(iLineOcc))
						Else
							sMasked1 = Left(sMasked1, oDictIgnoreTextColumn(iLineOcc)) & _
										string(oDictIgnoreTextLength(iLineOcc), "X") & _
										Right(sMasked1, len(sMasked1) - oDictIgnoreTextColumn(iLineOcc) - oDictIgnoreTextLength(iLineOcc))
						
							sMasked2 = Left(sMasked2, oDictIgnoreTextColumn(iLineOcc)) & _
										string(oDictIgnoreTextLength(iLineOcc), "X") & _
										Right(sMasked2, len(sMasked2) - oDictIgnoreTextColumn(iLineOcc) - oDictIgnoreTextLength(iLineOcc))
						End If
					Else
						sLog = sLog & "E: Ignore text specified beyond the end of the line at line " & iLine & vbLF
						bRetVal = False
						Exit Do
					End If
			End If
		Next

		'Now compare the lines
		If sMasked1 <> sMasked2 Then
			sLog = sLog & "E: Difference found at line " & iLine & vbLf & _
				"1st file line " & iLine & ":" & Replace(sMasked1, ", ", ",") & vbLf & _
				"2nd file line " & iLine & ":" & Replace(sMasked2, ", ", ",") & vbLf
			bRetVal = False
			Exit Do
		End If
		iLine = iLine + 1
	Loop

    Set oTextFile1 = Nothing
    Set oTextFile2 = Nothing
    Set oFs = Nothing

	TextFileCompare = bRetVal
End Function

'==============================================================================================
' Function/Sub: TextFileCompare2BAK(sTextFile1, sTextFile2, sIgnoreText) - return true if files match
' Purpose: 		sIgnoreText is in the form of row,column,length|row,column,length|row,column,length etc
'				sPatternList is in the form of Type,Pattern,MaxLength|Type,Pattern,MaxLength etc
'				where Type is TODAY and Pattern example is DD-MMM-YYYY (regular expressions can still be embedded)
'				and   Type is REGEX and Pattern is any regular expression
'==============================================================================================
Public Function TextFileCompare2BAK(sTextFile1, sTextFile2, sIgnoreText, sPatternList, sLog)
	Dim bRetVal
	Dim sLine1, sLine2, iLine, sMasked1, sMasked2
	Dim arrIgnoreText, oDictIgnoreTextColumn, oDictIgnoreTextLength
	Dim sEntry, arrEntry, iOccurrence, iLineOcc, sIgnoreText2
	Dim iCount1, iCount2 

	bRetVal = True

'****************************************************************
	'Read the files
	Dim oFS, oTextFile1, oTextFile2
    Set oFs = CreateObject("Scripting.FileSystemObject")
    Set oTextFile1 = oFs.OpenTextFile(sTextFile1, 1, False)
    Set oTextFile2 = oFs.OpenTextFile(sTextFile2, 1, False)

	Dim arrLines1, arrLines2
	arrLines1 = Split(oTextFile1.ReadAll, vbNewLine)
	arrLines2 = Split(oTextFile2.ReadAll, vbNewLine)
	
    Set oTextFile1 = Nothing
    Set oTextFile2 = Nothing
    Set oFs = Nothing
	
	'Check the line counts before removing blank lines
	If UBound(arrLines1) <> UBound(arrLines2) Then
		sLog = sLog & "I: Total number of lines are different for compared files. " & UBound(arrLines1)+1 & " vs " & UBound(arrLines2)+1 & vbLf
	Else
		sLog = sLog & "I: Total number of lines are the same for compared files. " & UBound(arrLines1)+1 & " vs " & UBound(arrLines2)+1 & vbLf
	End if
	
	'Remove blank lines
	Dim arrLineText1(), arrLineNums1(), dicNewToOld1
	Dim arrLineText2(), arrLineNums2(), dicNewToOld2
    Set dicNewToOld1 = CreateObject("Scripting.Dictionary")
    Set dicNewToOld2 = CreateObject("Scripting.Dictionary")
	Call RemoveBlankLines(arrLines1, arrLineText1, arrLineNums1, dicNewToOld1) 
	Call RemoveBlankLines(arrLines2, arrLineText2, arrLineNums2, dicNewToOld2)

	Dim sDebug, i, dicNewToOld2Keys, dicNewToOld2Items
	sDebug = ""
	dicNewToOld2Keys = dicNewToOld2.Keys
	dicNewToOld2items = dicNewToOld2.items
	For i = 0 to ubound(dicNewToOld2items)
		sDEbug = sdebug & dicNewToOld2Keys(i) & "=" & dicNewToOld2items(i) & ","
	Next
	msgbox sDEbug


	'Check the line counts after removing blank lines
	If UBound(arrLineText1) <> UBound(arrLineText2) Then
		sLog = sLog & "E: Total number of non-blank lines are different for compared files. " & UBound(arrLineText1)+1 & " vs " & UBound(arrLineText2)+1 & vbLf
		bRetVal = False
	Else
		sLog = sLog & "I: Total number of non-blank lines are the same for compared files. " & UBound(arrLineText1)+1 & " vs " & UBound(arrLineText2)+1 & vbLf
	End if
	
' ORig *********************************************************************
    Set oDictIgnoreTextColumn = CreateObject("Scripting.Dictionary")
    Set oDictIgnoreTextLength = CreateObject("Scripting.Dictionary")

	'Derive the ignore text list associated with the patternlist and add it to the ignoretext.
	'Use sTextFile2 because by convention that is the newly generated file with Today's dates
'	sIgnoreText = sIgnoreText & "|"& BuildIgnoreTextListFromArray(arrLineText2, sPatternList, "|", sLog)
	sIgnoreText = sIgnoreText & "|"& BuildIgnoreTextListFromFile(sTextFile2, sPatternList, "|", sLog)
	If sIgnoreText <> "|" Then	'Have some ignoring to do ...
		If mid(sIgnoreText,1,1) = "|" Then
			sIgnoreText = mid(sIgnoreText,2)
		End If

		If mid(sIgnoreText,len(sIgnoreText),1) = "|" Then
			sIgnoreText = mid(sIgnoreText,1,len(sIgnoreText)-1)
		End If

		arrIgnoreText = split(sIgnoreText, "|")
		For each sEntry in arrIgnoreText
			arrEntry = split(sEntry, ",")
			'Have to allow for multiple ignored text on a line - add 1000 to the line number each subsequent time it appears
			'Eg if the same line, say 2, has 5 bits of text to ignore then they will stored as 2, 1002, 2002, 3002 and 4002
			For iOccurrence = 0 to 10
				iLineOcc = (dicNewToOld2(cint(arrEntry(0)))) + iOccurrence*1000
				If  Not oDictIgnoreTextColumn.Exists(iLineOcc) Then Exit For
			Next
			
			'Check if the start and length to ignore are specified.
			If UBound(arrEntry) > 0 Then
				oDictIgnoreTextColumn.Add CInt(iLineOcc), CInt("0" & arrEntry(1))
				If UBound(arrEntry) > 1 Then
					oDictIgnoreTextLength.Add CInt(iLineOcc), CInt("0" & arrEntry(2))
				Else
					oDictIgnoreTextLength.Add CInt(iLineOcc), 0
				End If
			Else
				oDictIgnoreTextColumn.Add CInt(iLineOcc), 0
				oDictIgnoreTextLength.Add CInt(iLineOcc), 0
			End If

		Next
	End If
' ORig *********************************************************************
	
	iLine = 0
	Do While (iLine <= UBound(arrLineText1)) And (iLine <= UBound(arrLineText2)) 'Not oTextFile1.AtEndOfStream
		sLine1 = arrLineText1(iLine)
		sLine2 = arrLineText2(iLine)
		'Allow for the odd case where there is a comma without a following space.  Add a space after all commas. 
		sLine1 = Replace(sLine1, ",", ", ")
		sLine2 = Replace(sLine2, ",", ", ")
		'Mask the ignored text - equalise spaces and add @ to the start and end of each line to avoid checking for boundary conditions
		sMasked1 = "@" & EqualiseSpaces(sLine1, sLine2, iLine, 4, sLog) & "@"
		sMasked2 = "@" & sLine2 & "@"
		For iOccurrence = 0 to 10
			iLineOcc = dicNewToOld2(arrLineNums1(iLine)) + iOccurrence*1000
			If oDictIgnoreTextColumn.Exists(iLineOcc) Then
				'Check the line is long enough
					If len(sLine1) >= oDictIgnoreTextColumn(iLine)+oDictIgnoreTextLength(iLineOcc)-1  Then
						If oDictIgnoreTextLength(iLineOcc) = 0 Then
							'leave only the line content up to the column position
							sMasked1 = Left(sMasked1, oDictIgnoreTextColumn(iLineOcc))
							sMasked2 = Left(sMasked2, oDictIgnoreTextColumn(iLineOcc))
						Else
							sMasked1 = Left(sMasked1, oDictIgnoreTextColumn(iLineOcc)) & _
										string(oDictIgnoreTextLength(iLineOcc), "X") & _
										Right(sMasked1, len(sMasked1) - oDictIgnoreTextColumn(iLineOcc) - oDictIgnoreTextLength(iLineOcc))
						
							sMasked2 = Left(sMasked2, oDictIgnoreTextColumn(iLineOcc)) & _
										string(oDictIgnoreTextLength(iLineOcc), "X") & _
										Right(sMasked2, len(sMasked2) - oDictIgnoreTextColumn(iLineOcc) - oDictIgnoreTextLength(iLineOcc))
						End If
					Else
						sLog = sLog & "E: Ignore text specified beyond the end of the line at line " & arrLineNums1(iLine) & vbLF
						bRetVal = False
						Exit Do
					End If
			Else
				Exit for
			End If
		Next

		'Now compare the lines
		If sMasked1 <> sMasked2 Then
			sLog = sLog & "E: Difference found at line " & iLine & vbLf & _
				"1st file line " & arrLineNums1(iLine) & ":" & Replace(sMasked1, ", ", ",") & vbLf & _
				"2nd file line " & arrLineNums2(iLine) & ":" & Replace(sMasked2, ", ", ",") & vbLf
			bRetVal = False
			Exit Do
		End If
		iLine = iLine + 1
	Loop

	TextFileCompare2BAK = bRetVal
End Function

'==============================================================================================
' Function/Sub: TextFileCompare(sTextFile1, sTextFile2, sIgnoreText) - return true if files match
' Purpose: 		sIgnoreText is in the form of row,column,length|row,column,length|row,column,length etc
'				sPatternList is in the form of Type,Pattern,MaxLength|Type,Pattern,MaxLength etc
'				where Type is TODAY and Pattern example is DD-MMM-YYYY (regular expressions can still be embedded)
'				and   Type is REGEX and Pattern is any regular expression
'==============================================================================================
Public Function TextFileCompare2(sTextFile1, sTextFile2, sIgnoreText, sPatternList, sLog)
	Dim bRetVal
	Dim oFS, oTextFile1, oTextFile2
	Dim sLine1, sLine2, iLine, sMasked1, sMasked2
	Dim arrIgnoreText, oDictIgnoreTextColumn, oDictIgnoreTextLength
	Dim sEntry, arrEntry, iOccurrence, iLineOcc, sIgnoreText2
	Dim iCount1, iCount2 

	bRetVal = True

    'First check the line count of the files.
	iCount1 = CountFileLines(sTextFile1)
	iCount2 = CountFileLines(sTextFile2)
	If iCount1<> iCount2 Then
		sLog = sLog & "I: Line counts are different for compared files. " & iCount1 & " vs " & iCount2 & vbLf
	Else
		sLog = sLog & "I: Line counts are the same for compared files. " & iCount1 & " vs " & iCount2 & vbLf
	End if

    Dim sTextFile1NB, sTextFile2NB, arrNewToOld()
	ReDim arrNewToOld(iCount2)
    sTextFile1NB = RemoveBlankLines(sTextFile1,  arrNewToOld)
    sTextFile2NB = RemoveBlankLines(sTextFile2,  arrNewToOld)
    sIgnoreText = UpdateIgnoreText(sIgnoreText, arrNewToOld)

	'Check the line counts after removing blank lines
	iCount1 = CountFileLines(sTextFile1NB)
	iCount2 = CountFileLines(sTextFile2NB)
	If iCount1<> iCount2 Then
		sLog = sLog & "E: Total number of non-blank lines are different for compared files. " & iCount1 & " vs " & iCount2 & vbLf
		bRetVal = False
	Else
		sLog = sLog & "I: Total number of non-blank lines are the same for compared files. " & iCount1 & " vs " & iCount2 & vbLf
	End if


    Set oFs = CreateObject("Scripting.FileSystemObject")
    Set oTextFile1 = oFs.OpenTextFile(sTextFile1NB, 1, False)
    Set oTextFile2 = oFs.OpenTextFile(sTextFile2NB, 1, False)
	
    Set oDictIgnoreTextColumn = CreateObject("Scripting.Dictionary")
    Set oDictIgnoreTextLength = CreateObject("Scripting.Dictionary")

	'Derive the ignore text list associated with the patternlist and add it to the ignoretext.
	'Use sTextFile2 because by convention that is the newly generated file with Today's dates
	sIgnoreText = sIgnoreText & "|"& BuildIgnoreTextListFromFile2(sTextFile2NB, sPatternList, "|", arrNewToOld, sLog)
	If sIgnoreText <> "|" Then	'Have some ignoring to do ...
		If mid(sIgnoreText,1,1) = "|" Then
			sIgnoreText = mid(sIgnoreText,2)
		End If

		If mid(sIgnoreText,len(sIgnoreText),1) = "|" Then
			sIgnoreText = mid(sIgnoreText,1,len(sIgnoreText)-1)
		End If

		arrIgnoreText = split(sIgnoreText, "|")
		For each sEntry in arrIgnoreText
			arrEntry = split(sEntry, ",")
			'Have to allow for multiple ignored text on a line - add 1000 to the line number each subsequent time it appears
			'Eg if the same line, say 2, has 5 bits of text to ignore then they will stored as 2, 1002, 2002, 3002 and 4002
			For iOccurrence = 0 to 10
				iLineOcc = cint(arrEntry(0)) + iOccurrence*1000
				If  Not oDictIgnoreTextColumn.Exists(iLineOcc) Then Exit For
			Next
			
			'Check if the start and length to ignore are specified.
			If UBound(arrEntry) > 0 Then
				oDictIgnoreTextColumn.Add CInt(iLineOcc), CInt("0" & arrEntry(1))
				If UBound(arrEntry) > 1 Then
					oDictIgnoreTextLength.Add CInt(iLineOcc), CInt("0" & arrEntry(2))
				Else
					oDictIgnoreTextLength.Add CInt(iLineOcc), 0
				End If
			Else
				oDictIgnoreTextColumn.Add CInt(iLineOcc), 0
				oDictIgnoreTextLength.Add CInt(iLineOcc), 0
			End If

		Next
	End If
	
	iLine = 1
	Do While (iLine <= iCount1) And (iLine <= iCount2) 'Not oTextFile1.AtEndOfStream
		sLine1 = oTextFile1.ReadLine
		sLine2 = oTextFile2.ReadLine
		'Allow for the odd case where there is a comma without a following space.  Add a space after all commas. 
		sLine1 = Replace(sLine1, ",", ", ")
		sLine2 = Replace(sLine2, ",", ", ")
		'Mask the ignored text - equalise spaces and add @ to the start and end of each line to avoid checking for boundary conditions
		sMasked1 = "@" & EqualiseSpaces2(sLine1, sLine2, iLine, 4, arrNewToOld, sLog) & "@"
		sMasked2 = "@" & sLine2 & "@"
		For iOccurrence = 0 to 10
			iLineOcc = iLine + iOccurrence*1000
			If oDictIgnoreTextColumn.Exists(iLineOcc) Then
				'Check the line is long enough
					If len(sLine1) >= oDictIgnoreTextColumn(iLine)+oDictIgnoreTextLength(iLineOcc)-1  Then
						If oDictIgnoreTextLength(iLineOcc) = 0 Then
							'leave only the line content up to the column position
							sMasked1 = Left(sMasked1, oDictIgnoreTextColumn(iLineOcc))
							sMasked2 = Left(sMasked2, oDictIgnoreTextColumn(iLineOcc))
						Else
							sMasked1 = Left(sMasked1, oDictIgnoreTextColumn(iLineOcc)) & _
										string(oDictIgnoreTextLength(iLineOcc), "X") & _
										Right(sMasked1, len(sMasked1) - oDictIgnoreTextColumn(iLineOcc) - oDictIgnoreTextLength(iLineOcc))
						
							sMasked2 = Left(sMasked2, oDictIgnoreTextColumn(iLineOcc)) & _
										string(oDictIgnoreTextLength(iLineOcc), "X") & _
										Right(sMasked2, len(sMasked2) - oDictIgnoreTextColumn(iLineOcc) - oDictIgnoreTextLength(iLineOcc))
						End If
					Else
						sLog = sLog & "E: Ignore text specified beyond the end of the line at line " & arrNewToOld(iLine) & "(" & iLine & ")" & vbL
						bRetVal = False
						Exit Do
					End If
			End If
		Next

		'Now compare the lines
		If sMasked1 <> sMasked2 Then
			sLog = sLog & "E: Difference found at line " & arrNewToOld(iLine) & "(" & iLine & ")" & vbLf & _
				"1st file line " & arrNewToOld(iLine) & "(" & iLine & ")" & ":" & Replace(sMasked1, ", ", ",") & vbLf & _
				"2nd file line " & arrNewToOld(iLine) & "(" & iLine & ")" & ":" & Replace(sMasked2, ", ", ",") & vbLf
			bRetVal = False
			Exit Do
		End If
		iLine = iLine + 1
	Loop

	oTextFile1.Close
	oTextFile2.Close
    Set oTextFile1 = Nothing
    Set oTextFile2 = Nothing
    
    oFS.DeleteFile sTextFile1NB, True
    oFS.DeleteFile sTextFile2NB, True
    Set oFs = Nothing

	TextFileCompare2 = bRetVal
End Function

'==============================================================================================
' Function/Sub: RemoveBlankLines(sInfile, byref dicNewToOld))
' Purpose: 	
'==============================================================================================
Function RemoveBlankLines(sInfile, byref arrNewToOld)
	Dim iLineNum, sLine
	Dim iInLIne, iOutLIne

	'Read the files
	Dim sOutfile
	Dim oFS, oInFile, oOutfile
	sOutfile = sInFile & ".NoBlankS"
    Set oFs = CreateObject("Scripting.FileSystemObject")
    Set oInFile = oFs.OpenTextFile(sInfile, 1, False)
    Set oOutfile = oFs.OpenTextFile(sOutfile, 2, True)

	iInLIne = 1 : iOutLIne = 1
	Do While Not oInFile.AtEndOfStream
		sLIne = oInFile.ReadLine
		If Trim(sline) <> "" Then
			oOutfile.WriteLine sLine
			arrNewToOld(iOutLIne) = iInLIne
			iOutLIne = iOutLIne + 1
		End If
		iInLIne = iInLIne+1
	Loop
	
	oOutfile.Close
	oInFile.Close

    Set oOutfile = Nothing
    Set oInFile = Nothing
    Set oFs = Nothing
	
	ReDim Preserve arrNewToOld(iOutLIne)

	RemoveBlankLines = sOutfile

End Function

'==============================================================================================
' Function/Sub: UpdateIgnoreText(sIgnoreText, arrNewToOld)
' Purpose: 	
'==============================================================================================
Function UpdateIgnoreText(sIgnoreText, arrNewToOld)
	Dim sUpdatedIgnoreText
	Dim arrIgnoreText, sEntry, arrEntry, iEntry
	Dim dicOldToNew, i

	If sIgnoreText <> "" Then 
	    Set dicOldToNew = CreateObject("Scripting.Dictionary")
	    For i = 1 To ubound(arrNewToOld)
	    	dicOldToNew.Add arrNewToOld(i), i
	    Next

		sUpdatedIgnoreText = ""
		arrIgnoreText = split(sIgnoreText, "|")
		For each sEntry in arrIgnoreText
			arrEntry = split(sEntry, ",")
			sUpdatedIgnoreText = sUpdatedIgnoreText & dicOldToNew(cint(arrEntry(0)))	'replace with a new line number
			If UBound(arrEntry) > 0 Then
				For i = 1 To UBound(arrEntry)
					sUpdatedIgnoreText = sUpdatedIgnoreText & "," & arrEntry(i)		'Add the ",start,length"
				Next
			End If
			sUpdatedIgnoreText = sUpdatedIgnoreText & "|"
		Next
		sUpdatedIgnoreText = Left(sUpdatedIgnoreText, Len(sUpdatedIgnoreText)-1)	'KNock of the last |
		
		Set dicOldToNew = Nothing
	Else
		sUpdatedIgnoreText = sIgnoreText
	End if
	
	UpdateIgnoreText = sUpdatedIgnoreText
End Function

'==============================================================================================
' Function/Sub: EqualiseSpaces()
' Purpose: 	Equalise the white spaces in string1 so they match the white spaces in string2 
' 			provided their difference is within the given tolerance
'==============================================================================================
Function EqualiseSpaces(sString1, sString2, iLine, iTolerance, sLog)
	Dim sRet, iCount, iStart
	Dim oRegExp, oMatches, oMatch
	Dim arrSpacesStart1, arrSpacesLength1
	Dim arrSpacesStart2, arrSpacesLength2
	Dim sTemp1, sTemp2

	'Add an "@ " to the beginning and " @" to the end so that there are no leading or trailing spaces
	sTemp1 = "@ " & sString1 & " @"
	sTemp2 = "@ " & sString2 & " @"

	Set oRegExp = New RegExp		' Create a regular expression.
	oRegExp.Pattern = " +"			' One or more spaces.
	oRegExp.IgnoreCase = True       ' Set case insensitivity.
	oRegExp.Global = True         	' Set global applicability.

	'Get interior spaces for String 1
	iCount = 0
	arrSpacesStart1 = Array()
	arrSpacesLength1 = Array()
	Set oMatches = oRegExp.Execute(sTemp1)	' Execute search.
	For Each oMatch in oMatches      			' Iterate Matches collection.
		ReDim Preserve arrSpacesStart1(iCount)
		ReDim Preserve arrSpacesLength1(iCount)
		arrSpacesStart1(iCount) = oMatch.FirstIndex + 1
		arrSpacesLength1(iCount) = oMatch.Length
		iCount = iCount + 1
	Next
	Set oMatch = Nothing
	Set oMatches = Nothing

	'Get interior spaces for String 2
	iCount = 0
	arrSpacesStart2 = Array()
	arrSpacesLength2 = Array()
	Set oMatches = oRegExp.Execute(sTemp2)	' Execute search.
	For Each oMatch in oMatches      			' Iterate Matches collection.
		ReDim Preserve arrSpacesStart2(iCount)
		ReDim Preserve arrSpacesLength2(iCount)
		arrSpacesStart2(iCount) = oMatch.FirstIndex + 1
		arrSpacesLength2(iCount) = oMatch.Length
		iCount = iCount + 1
	Next
	Set oMatch = Nothing
	Set oMatches = Nothing
	Set oRegExp = Nothing

	'Are there the same number of interior sets of spaces?
	If UBound(arrSpacesStart1) <>  UBound(arrSpacesStart2) Then
		sRet = sString1	' Different number of white spaces. Just return the original string1
		sLog = sLog & "W: Line " & iLine & " - Different number of white spaces." & vbLF & sString1 & vbLf & sString2 & vbLF
		EqualiseSpaces = sRet : Exit Function
	Else	'Now rebuild String1 while equalising spaces
		sRet = ""
		iStart = 1
		For iCount = 0 to uBound(arrSpacesStart1)
			If arrSpacesLength1(iCount) = arrSpacesLength2(iCount) Then 
				sRet = sRet & Mid(sTemp1, iStart, arrSpacesStart1(iCount)-iStart) & Mid(sTemp2, arrSpacesStart2(iCount), arrSpacesLength2(iCount))
				iStart = arrSpacesStart1(iCount) + arrSpacesLength1(iCount)
			ElseIf abs(arrSpacesLength1(iCount) - arrSpacesLength2(iCount)) <= iTolerance Then
				sRet = sRet & Mid(sTemp1, iStart, arrSpacesStart1(iCount)-iStart) & Mid(sTemp2, arrSpacesStart2(iCount), arrSpacesLength2(iCount))
				iStart = arrSpacesStart1(iCount) + arrSpacesLength1(iCount)
				sLog = sLog & "W: Line " & iLine & " - White space length different but within acceptable tolerance." & vbLF & sString1 & vbLf & sString2 & vbLF
			Else	'Outside tolerances so return the original string1 
				sRet = sString1	
				sLog = sLog & "W: Line " & iLine & " - White space length difference exceeds acceptable tolerance." & vbLF & sString1 & vbLf & sString2 & vbLF
				EqualiseSpaces = sRet : Exit Function
			End If
		Next
		
		sRet = sRet & Mid(sTemp1, iStart)
	End If

	EqualiseSpaces = Mid(sRet, 3, len(sRet)-4) 'Remove the "@ " and " @"

End Function

'==============================================================================================
' Function/Sub: EqualiseSpaces2()
' Purpose: 	Equalise the white spaces in string1 so they match the white spaces in string2 
' 			provided their difference is within the given tolerance
'==============================================================================================
Function EqualiseSpaces2(sString1, sString2, iLine, iTolerance, arrNewToOld, sLog)
	Dim sRet, iCount, iStart
	Dim oRegExp, oMatches, oMatch
	Dim arrSpacesStart1, arrSpacesLength1
	Dim arrSpacesStart2, arrSpacesLength2
	Dim sTemp1, sTemp2

	'Add an "@ " to the beginning and " @" to the end so that there are no leading or trailing spaces
	sTemp1 = "@ " & sString1 & " @"
	sTemp2 = "@ " & sString2 & " @"

	Set oRegExp = New RegExp		' Create a regular expression.
	oRegExp.Pattern = " +"			' One or more spaces.
	oRegExp.IgnoreCase = True       ' Set case insensitivity.
	oRegExp.Global = True         	' Set global applicability.

	'Get interior spaces for String 1
	iCount = 0
	arrSpacesStart1 = Array()
	arrSpacesLength1 = Array()
	Set oMatches = oRegExp.Execute(sTemp1)	' Execute search.
	For Each oMatch in oMatches      			' Iterate Matches collection.
		ReDim Preserve arrSpacesStart1(iCount)
		ReDim Preserve arrSpacesLength1(iCount)
		arrSpacesStart1(iCount) = oMatch.FirstIndex + 1
		arrSpacesLength1(iCount) = oMatch.Length
		iCount = iCount + 1
	Next
	Set oMatch = Nothing
	Set oMatches = Nothing

	'Get interior spaces for String 2
	iCount = 0
	arrSpacesStart2 = Array()
	arrSpacesLength2 = Array()
	Set oMatches = oRegExp.Execute(sTemp2)	' Execute search.
	For Each oMatch in oMatches      			' Iterate Matches collection.
		ReDim Preserve arrSpacesStart2(iCount)
		ReDim Preserve arrSpacesLength2(iCount)
		arrSpacesStart2(iCount) = oMatch.FirstIndex + 1
		arrSpacesLength2(iCount) = oMatch.Length
		iCount = iCount + 1
	Next
	Set oMatch = Nothing
	Set oMatches = Nothing
	Set oRegExp = Nothing

	'Are there the same number of interior sets of spaces?
	If UBound(arrSpacesStart1) <>  UBound(arrSpacesStart2) Then
		sRet = sString1	' Different number of white spaces. Just return the original string1
		sLog = sLog & "W: Line " & arrNewToOld(iLine) & "(" & iLine & ")" & " - Different number of white spaces." & vbLF & sString1 & vbLf & sString2 & vbLF
		EqualiseSpaces2 = sRet : Exit Function
	Else	'Now rebuild String1 while equalising spaces
		sRet = ""
		iStart = 1
		For iCount = 0 to uBound(arrSpacesStart1)
			If arrSpacesLength1(iCount) = arrSpacesLength2(iCount) Then 
				sRet = sRet & Mid(sTemp1, iStart, arrSpacesStart1(iCount)-iStart) & Mid(sTemp2, arrSpacesStart2(iCount), arrSpacesLength2(iCount))
				iStart = arrSpacesStart1(iCount) + arrSpacesLength1(iCount)
			ElseIf abs(arrSpacesLength1(iCount) - arrSpacesLength2(iCount)) <= iTolerance Then
				sRet = sRet & Mid(sTemp1, iStart, arrSpacesStart1(iCount)-iStart) & Mid(sTemp2, arrSpacesStart2(iCount), arrSpacesLength2(iCount))
				iStart = arrSpacesStart1(iCount) + arrSpacesLength1(iCount)
				sLog = sLog & "W: Line " & arrNewToOld(iLine) & "(" & iLine & ")" & " - White space length different but within acceptable tolerance." & vbLF & sString1 & vbLf & sString2 & vbLF
			Else	'Outside tolerances so return the original string1 
				sRet = sString1	
				sLog = sLog & "W: Line " & arrNewToOld(iLine) & "(" & iLine & ")" & " - White space length difference exceeds acceptable tolerance." & vbLF & sString1 & vbLf & sString2 & vbLF
				EqualiseSpaces2 = sRet : Exit Function
			End If
		Next
		
		sRet = sRet & Mid(sTemp1, iStart)
	End If

	EqualiseSpaces2 = Mid(sRet, 3, len(sRet)-4) 'Remove the "@ " and " @"

End Function
'===================================================================================
' Function: DeleteZipFile(sZipFile)
'===================================================================================
Sub DeleteZipFile(sZipFile, sLog)
	Dim oFS

	Set oFS = CreateObject("Scripting.FileSystemObject")
	'Does file to add exist?
	If oFS.FileExists(sZipFile) Then 
		oFS.DeleteFile sZipFile, True
		sLog = sLog & "D: File " & sZipFile & " deleted." &vbLf
	Else
		sLog = sLog & "W: File to delete doesn't exist: " & sZipFile & vbLf
	End If

	Set oFS = Nothing
End sub

'===================================================================================
' Function: AddFoldertoZip(sFilename, sZipFile)
'===================================================================================
Function AddFoldertoZip(sFilename, sZipFile, sLog)
	Dim iCount, iTimeout, iTemp
	Dim oApp, oFS, oFile, oZip
	Const ForWriting = 2
	
	iTimeout = 30	'Seconds
	Set oFS = CreateObject("Scripting.FileSystemObject")
	'Does file to add exist?
	If oFS.FolderExists(sFilename) Then
		'Ceate zip file?
		If Not oFS.FileExists(sZipFile) Then
			Set oZip = oFS.OpenTextFile(sZipFIle, ForWriting, True)
			oZip.Write "PK" & Chr(5) & Chr(6) & String(18, Chr(0))
			oZip.Close
			Set oZip = Nothing
			sLog = sLog & "D: Created zip file " & sZipFile & vblf
		End If
	Else
		sLog = sLog & "D: File " & sFilename & " not found." & vblf
		AddFoldertoZip = False : Exit Function
	End If

	'Add the file to the zip file
	'Set oFile = oFS.GetFile(sFilename)
    ' Create a Shell object
    Set oApp = CreateObject("Shell.Application")
    ' Copy the files to the compressed folder
    iCount = oApp.NameSpace(sZipFile).Items.Count
	
	oApp.NameSpace(sZipFile).CopyHere sFilename
	
	'WaitTime 0,200
	WaitTime(200)
	'WaitTime until the file is ready, otherwise seeing "oApp.NameSpace(sZipFile)" error
	On Error Resume Next
	iTemp = oApp.NameSpace(sZipFile).Items.Count
	Do While Err.Number <> 0
		Err.Clear
		'WaitTime 1
		WaitTime(100)
		iTemp = oApp.NameSpace(sZipFile).Items.Count
	Loop
	Err.Clear
	On Error Goto 0
	
	' Keep script waiting until compression is done
    Do Until oApp.NameSpace(sZipFile).Items.Count = iCount + 1 or iTimeout <= 0
        'WScript.Sleep 1000
		'WaitTime 1
		WaitTime(100)
		iTimeout = iTimeout - 1
    Loop

	sLog = sLog & "D: Added " & sFileName & " to zip file " & sZipFile & vbLf
	
	'Set oFile = Nothing
	Set oApp = Nothing
	Set oFS = Nothing

	AddFoldertoZip = True
End Function

'===================================================================================
' Function: AddFiletoZip(sFilename, sZipFile)
'===================================================================================
Function AddFiletoZip(sFilename, sZipFile, sLog)
	Dim iCount, iTimeout, iTemp
	Dim oApp, oFS, oFile, oZip
	Const ForWriting = 2
	
	iTimeout = 30	'Seconds
	Set oFS = CreateObject("Scripting.FileSystemObject")
	'Does file to add exist?
	If oFS.FileExists(sFilename) Then
		'Ceate zip file?
		If Not oFS.FileExists(sZipFile) Then
			Set oZip = oFS.OpenTextFile(sZipFIle, ForWriting, True)
			oZip.Write "PK" & Chr(5) & Chr(6) & String(18, Chr(0))
			oZip.Close
			Set oZip = Nothing
			sLog = sLog & "D: Created zip file " & sZipFile & vblf
		End If
	Else
		sLog = sLog & "D: File " & sFilename & " not found." & vblf
		AddFiletoZip = False : Exit Function
	End If

	'Add the file to the zip file
	'Set oFile = oFS.GetFile(sFilename)
    ' Create a Shell object
    Set oApp = CreateObject("Shell.Application")
    ' Copy the files to the compressed folder
    iCount = oApp.NameSpace(sZipFile).Items.Count
	
	oApp.NameSpace(sZipFile).CopyHere sFilename
	
	'WaitTime 0,200
	WaitTime(200)

	'WaitTime until the file is ready, otherwise seeing "oApp.NameSpace(sZipFile)" error
	On Error Resume Next
	iTemp = oApp.NameSpace(sZipFile).Items.Count
	Do While Err.Number <> 0
		Err.Clear
		'WaitTime 1
		WaitTime(100)
		iTemp = oApp.NameSpace(sZipFile).Items.Count
	Loop
	Err.Clear
	On Error Goto 0
	
	' Keep script waiting until compression is done
    Do Until oApp.NameSpace(sZipFile).Items.Count = iCount + 1 or iTimeout <= 0
        'WScript.Sleep 1000
		'WaitTime 1
		WaitTime(100)
		iTimeout = iTimeout - 1
    Loop

	sLog = sLog & "D: Added " & sFileName & " to zip file " & sZipFile & vbLf
	
	'Set oFile = Nothing
	Set oApp = Nothing
	Set oFS = Nothing

	AddFiletoZip = True
End Function

Public Function ZIPFolder(sFilepath, sZipExecPath, sZipOutput, sLog)
	Dim oWshScriptExec, oFS, oShell, oStdOut
	Dim sLastLine
	Set oShell = CreateObject("WScript.Shell")
	Set oFS = CreateObject("Scripting.FileSystemObject")
	
	if oFS.FolderExists(sFilepath) Then
		Set oWshScriptExec = oShell.Exec(sZipExecPath & " a -tzip -r " & sZipOutput & " " & sFilepath)
		Set oStdOut = oWshScriptExec.StdOut
		While Not oStdOut.AtEndOfStream
		   sLastLine = oStdOut.ReadLine
		   sLog = sLog & vbCrLf & sLastLine
		Wend
		If sLastLine = "Everything is Ok" Then
			sLog = sLog & "E: Folder successfuly packed"
			ZIPFolder = True
		Else
			sLog = sLog & "E: Zip doesn't exist"
			ZIPFolder = False
		End If
	Else
		sLog = sLog & "E: Folder doesn't exist"
		ZIPFolder = False
	end if	
End Function

'===================================================================================
' Function: EscapeString4Regex(sString)
'===================================================================================
Function EscapeString4Regex(sString)
	Dim sRetVal, i
	Dim sRegexChars, sRegexChar
		
	sRegexChars ="\^$*+?.()|{}[]"
	sRetVal = sString

	For i = 1 To Len(sRegexChars)
		sRegexChar = Mid(sRegexChars, i, 1)
		sRetVal = Replace(sRetVal, sRegexChar, "\" & sRegexChar)
	Next
	
	EscapeString4Regex = sRetVal

End Function
'===================================================================================
' Function: AddBlankLinesToFile(sTextFile, iWhere, iNumber)
'===================================================================================
Public Sub AddBlankLinesToFile(sTextFile, iWhere, iNumber)
	Dim oFS, oFile
	Dim arrLines, iWhereActual, iCount, i
	
	Set oFS = CreateObject("Scripting.FileSystemObject")
	Set oFile = oFS.OpenTextFile(sTextFile, 1, False)
	arrLines = Split(oFile.ReadAll, vbNewLine)
	oFile.Close
	Set oFile = Nothing

	iCount = CountFileLines(sTextFile)

	Set oFile = oFS.OpenTextFile(sTextFile, 2, True)

	iWhereActual = iWhere
	If iCount < iWhere Then iWhereActual = Icount + 1
	
	For i = 0 To iWhereActual-1
		oFile.WriteLine arrLines(i)
	Next

	oFile.WriteBlankLines iNumber	
	For i = iWhere To UBound(arrLines)
		oFile.WriteLine arrLines(i)
	Next
	
	oFile.Close
	Set oFile = Nothing
	Set oFS = Nothing
End Sub

'==============================================================================================
' Function/Sub: CheckBuildPath(sPath)
' Purpose: Check/build (if does not exist) path.
'==============================================================================================
Public Sub CheckBuildPath(sPath)
	Dim oFS, aPath, i

	Set oFS = CreateObject("Scripting.FileSystemObject")		

	aPath = Split(sPath, "\")
	sPath = aPath(0)
	
	For i=1 to UBound(aPath)
		sPath = sPath & "\" & aPath(i)
		If (Not oFS.FolderExists(sPath)) Then
			oFS.CreateFolder(sPath)
		End if
	Next
	
	Set oFS = Nothing
End Sub

Public Function DMLY_NOSLASH(sDate, sLocale)
	Dim dDate
	dDate = CDate(sDate)
	Select Case sLocale
		Case "U.S."
			DMLY_NOSLASH = Right("0" & CStr(Day(dDate)), 2) & "-" & CStr(MonthName(Month(dDate), True)) & "-" & Year(dDate) 				
		Case "Europe"
			DMLY_NOSLASH = Right("0" & CStr(Day(dDate)), 2) & "-" & CStr(MonthName(Month(dDate), True)) & "-" & Year(dDate) 		
	End Select	
End Function

Public Function DMY_NOSLASH(sDate, sLocale)
	Dim dDate
	dDate = CDate(sDate)
	Select Case sLocale
		Case "U.S."
			DMY_NOSLASH = Right("0" & CStr(Day(dDate)), 2) & "-" & CStr(MonthName(Month(dDate), True)) & "-" & Right(Year(dDate), 2)
		Case "Europe"
			DMY_NOSLASH = Right("0" & CStr(Day(dDate)), 2) & "-" & CStr(MonthName(Month(dDate), True)) & "-" & Right(Year(dDate), 2)
	End Select  	
End Function

Public Function MDSY_SLASH(sDate, sLocale)
	Dim dDate
	dDate = CDate(sDate)	
	Select Case sLocale
		Case "U.S."
			MDSY_SLASH = Right("0" & CStr(Month(dDate)), 2) & "/" & Right("0" & CStr(Day(dDate)), 2) & "/" & Right(Year(dDate), 2)
		Case "Europe"
			MDSY_SLASH = Right("0" & CStr(Day(dDate)), 2) & "/" & Right("0" & CStr(Month(dDate)), 2) & "/" & Right(Year(dDate), 2)
	End Select          	
End Function

Public Function MDY_PERIOD(sDate, sLocale)
	Dim dDate
	dDate = CDate(sDate)	
		Select Case sLocale
		Case "U.S."
			MDY_PERIOD = Right("0" & CStr(Month(dDate)), 2) & "." & Right("0" & CStr(Day(dDate)), 2) & "." & Year(dDate)
		Case "Europe"
			MDY_PERIOD = Right("0" & CStr(Day(dDate)), 2) & "." & Right("0" & CStr(Month(dDate)), 2) & "." & Year(dDate)
	End Select   
End Function

Public Function MDY_SLASH(sDate, sLocale)
	Dim dDate
	dDate = CDate(sDate)	
	Select Case sLocale
		Case "U.S."
			MDY_SLASH = Right("0" & CStr(Month(dDate)), 2) & "/" & Right("0" & CStr(Day(dDate)), 2) & "/" & Year(dDate)
		Case "Europe"
			MDY_SLASH = Right("0" & CStr(Day(dDate)), 2) & "/" & Right("0" & CStr(Month(dDate)), 2) & "/" & Year(dDate)
	End Select   
End Function

Public Function MINIMAL(sDate, sLocale)
	Dim dDate
	dDate = CDate(sDate)	
	Select Case sLocale
		Case "U.S."
			MINIMAL = Right("0" & CStr(Day(dDate)), 2) & CStr(MonthName(Month(dDate), True)) & Right(Year(dDate), 2)
		Case "Europe"
			MINIMAL = Right("0" & CStr(Day(dDate)), 2) & CStr(MonthName(Month(dDate), True)) & Right(Year(dDate), 2)
	End Select   	
End Function

Public Function ISO8601_EXTENDED(sDate, sLocale)
	Dim dDate
	dDate = CDate(sDate)	
	Select Case sLocale
		Case "U.S."
			ISO8601_EXTENDED = Year(dDate) & "-" & Right("0" & CStr(Month(dDate)), 2) & "-" & Right("0" & CStr(Day(dDate)), 2) 
		Case "Europe"
			ISO8601_EXTENDED = Year(dDate) & "-" & Right("0" & CStr(Month(dDate)), 2) & "-" & Right("0" & CStr(Day(dDate)), 2) 
	End Select      	
End Function

Public Function ISO8601(sDate, sLocale)
	Dim dDate
	dDate = CDate(sDate)	
	Select Case sLocale
		Case "U.S."
			ISO8601 = Year(dDate) & Right("0" & CStr(Month(dDate)), 2) & Right("0" & CStr(Day(dDate)), 2) 
		Case "Europe"
			ISO8601 = Year(dDate) & Right("0" & CStr(Month(dDate)), 2) & Right("0" & CStr(Day(dDate)), 2) 
	End Select      	
End Function
''===================================================================================
'===================================================================================
'Function : CheckFolderExists
'Description : Checks whether the folder exists in the path and returns True or False
'===================================================================================
Public Function CheckFolderExists(sPath)
	Dim oFSO, sExists
	
	Set oFSO = CreateObject("Scripting.FileSystemObject")
		
	'Check the path exists
	If (oFSO.FolderExists(sPath)) Then
		sExists = True
	Else
		sExists = False
	End If

	Set oFSO = Nothing
	CheckFolderExists = sExists
End Function
'===================================================================================
'Function : CheckFileExists
'Description : Checks whether the folder exists and if true then checks whether file exists in the path and 
'                     returns True or False
'===================================================================================
Public Function CheckFileExists(sPath, sFileName)
	Dim oFSO, sExists
	
	Set oFSO = CreateObject("Scripting.FileSystemObject")

	'Check the path exists
	If (oFSO.FolderExists(sPath)) Then
		'Check the file exists
		If (oFSO.FileExists(sPath & sFileName)) Then
			sExists = True
		Else
			sExists = False
		End If
	Else
		sExists = False
	End If

    Set oFSO = Nothing
	CheckFileExists = sExists
End Function

'===================================================================================
'Function : CreateFolder
'Description : Checks whether the folder exists and if not it creates it.
'				does this recursively for its parent(s)
'===================================================================================
Public Sub CreateFolder(sFolder)
	Dim oFS

	Set oFS = CreateObject("Scripting.FileSystemObject")

	If not oFS.FolderExists(sFolder) Then
		If not oFS.FolderExists(oFS.GetParentFolderName(sFolder)) Then CreateFolder(oFS.GetParentFolderName(sFolder))
		oFS.CreateFolder(sFolder)
	End If

	Set oFS = Nothing
End Sub

'===================================================================================
'Function : FolderExists
'Description : True if the folder exists.
'===================================================================================
Public Function FolderExists(sFolder)
	Dim oFS, bRetVal
	
	bRetVal = False

	Set oFS = CreateObject("Scripting.FileSystemObject")
	If oFS.FolderExists(sFolder) Then bRetVal = True
	Set oFS = Nothing
	
	FolderExists = bRetVal

End Function

'==============================================================================================
' Function/Sub: CopyAttachmentToFS(sPrefix, sExtension, sDestinationFolder)
' Purpose: Copies attachment from QC and returns creation date and time of the file
'==============================================================================================
Function CopyAttachmentToFS(sPrefix, sExtension, sDestinationFolder)
'	Dim oQC,
	Dim oFS
	Dim sFilePath
	
'	Set oQC = New clsQC
	Set oFS = CreateObject("Scripting.FileSystemObject")
	
	'download attachment and copy it to destination folder
	If oQC.IsQCRun() = True Then
		If FolderExists(sDestinationFolder) Then
			sFilePath = oQC.GetAttachmentFileFromQC(sPrefix, sExtension, sDestinationFolder)
		End If	
	End If	
	
	CopyAttachmentToFS = oFS.GetFile(sFilePath).DateCreated
End Function

'==============================================================================================
' Function/Sub: IsRegexMatching(sPattern, sString, bIgnoreCase)
' Purpose: IsRegexMatching returns true, if given string matches given pattern
'==============================================================================================
Function IsRegexMatching(sPattern, sString, bIgnoreCase)
	Dim oRegEx
	
    Set oRegEx = New RegExp	
    
	oRegEx.Pattern = sPattern
	oRegEx.IgnoreCase = bIgnoreCase
	
	IsRegexMatching = oRegEx.Test(sString)
	
	Set oRegEx = Nothing
End Function

Function GetRegexMatches(ByRef arrResult, sPattern, sString, bIgnoreCase)
	Dim oRegEx, oMatches, oMatch
	Dim arrRes()
	Dim i, iCount
	
    Set oRegEx = New RegExp	
	
	oRegEx.Pattern = sPattern
	oRegEx.IgnoreCase = bIgnoreCase
	oRegEx.Global = True
	Set oMatches = oRegEx.Execute(sString)
	iCount = oMatches.Count
	If iCount > 0 Then
		ReDim arrRes(iCount - 1)
		i = 0
		For each oMatch  in oMatches
			arrRes(i) = oMatch.Value
			i = i + 1
		Next
		arrResult = arrRes
	End If
	GetRegexMatches = iCount
	
	Set oRegEx = Nothing
End Function

Private Function LocalFileExists(sFilePath)
	Dim oFS
	Set oFS = CreateObject("Scripting.FileSystemObject") 
	LocalFileExists = oFS.FileExists(sFilePath)
	Set oFS = Nothing
End Function

Private Function DeleteFile(sFilePath)
	Dim oFS
	Set oFS = CreateObject("Scripting.FileSystemObject") 
	If oFS.FileExists(sFilePath) Then
		oFS.DeleteFile sFilePath, True
	End If
	Set oFS = Nothing
End Function

Private Function FileRename(sFilePath, sNewName)
	Dim oFS, oFile
	Set oFS = CreateObject("Scripting.FileSystemObject") 
	If oFS.FileExists(sFilePath) Then
		Set oFile = oFS.GetFile(sFilePath)
		oFile.Name = sNewName
		FileRename = oFile.Path
		Set oFile = Nothing
	End If
	Set oFS = Nothing
End Function

Function DictionaryStringPreprocessing(sString, sDelimiter, sSeparator)
   Dim i
   Dim arrValues
	arrValues = Split(sString, sDelimiter)
	For i = 0 To UBound(arrValues)
		If UBound(Split(arrValues(i), sSeparator)) = 0 Then
			arrValues(i) = arrValues(i) & sSeparator & arrValues(i)
		End If
	Next
	
	DictionaryStringPreprocessing = Join(arrValues, sDelimiter)
End Function
'==============================================================================================
' Function/Sub: WaitTime(iTimeToWait)
' Purpose: function that will stop script for specific period of time
' Param: iTimeToWait = time to WaitTime in miliseconds
'==============================================================================================
Function WaitTime(iTimeToWait)
	On Error Resume Next
		Select Case globRunMode
			Case QTP_TEST,QTP_LOCAL_TEST
				Wait(iTimeToWait/1000)		
			Case VAPI_XP_TEST
				XTools.Sleep(iTimeToWait)	
			Case CMD_TEST
				WScript.Sleep(iTimeToWait)
		End Select		 
	If Err.Number <> 0 Then			
		WaitTime = XL_DISPATCH_FAIL
		oVBSFramework.oTraceLog.Message(Array("Error in WaitTime function",LOG_WARNING))	
		Exit function
	End If
	WaitTime = XL_DISPATCH_PASS	
End Function
'==============================================================================================
' Function/Sub: WaitTime(iSec,iMiliSec)
' Purpose: function that will stop script for specific period of time
' Param: iTimeToWait = time to WaitTime in miliseconds
'==============================================================================================
Function WaitTime2(iSec, iMiliSec)
	On Error Resume Next
		Select Case globRunMode
			Case QTP_TEST,QTP_LOCAL_TEST
				Wait iSec, iMiliSec		
			Case VAPI_XP_TEST
				XTools.Sleep((iSec*1000) + iMiliSec)	
			Case CMD_TEST
				WScript.Sleep((iSec*1000) + iMiliSec)
		End Select		 
	If Err.Number <> 0 Then			
		WaitTime = XL_DISPATCH_FAIL
		oVBSFramework.oTraceLog.Message(Array("Error in WaitTime(iSec,iMiliSec) function",LOG_WARNING))	
		Exit function
	End If
	WaitTime = XL_DISPATCH_PASS	
End Function

Private sRecipients, sSubject, sBody, iWait
'==============================================================================================
' Function/Sub: SendMailMail(oRow)
' Purpose:
'	Send mail specified by params
' Parameters:
'	recipient, subject, body
' Returns: sending time,sending date
'==============================================================================================
Function SendMail(oRow, sLog)
	'Dim sRecipients, sSubject, sBody, iWait
	Dim oApp, oNs, msg, Recpt, oFolder, item1

	Const sRoutine = "clsFaxCheck.SendMail"
	oVBSFramework.oTraceLog.Entered(sRoutine)
	
	'parameter handling
	sRecipients = CStr(oRow.Cells(1, XL_PARM_001).Value)
	sSubject = CStr(oRow.Cells(1, XL_PARM_002).Value)
	sBody = CStr(oRow.Cells(1, XL_PARM_003).Value)
	iWait = CStr(oRow.Cells(1, XL_PARM_004).Value)
	
	SendMail = Send(oRow, sRecipients, sSubject, sBody, iWait, sLog)
	oVBSFramework.oTraceLog.Exited(sRoutine)
End Function

'==============================================================================================
' Function/Sub: Send(oRow)
' Purpose: Called from SendMessage
' Parameters:
' Returns:
'==============================================================================================
Private Function Send(oRow, sRecipients, sSubject, sBody, iWait, sLog)
	'Dim sRecipients, sSubject, sBody, iWait
	Dim oApp, oNs, msg, Recpt, oFolder, item1', dTime

	Const sRoutine = "clsFaxCheck.Send"
	oVBSFramework.oTraceLog.Entered(sRoutine)
	If iWait = "" or Not IsNumeric(iWait) Then 
		iWait = 10
	End If

	Set oApp = CreateObject("Outlook.Application")
	WaitTime(iWait)
	Set oNs = oApp.GetNameSpace("MAPI")
	If oNs Is Nothing Then
		sLog = sLog & "E: Error launching Outlook Application" & vbLf
		Send = XL_DISPATCH_FAIL : oVBSFramework.oTraceLog.Exited(sRoutine) : Exit Function
	End If
	
	
	oNs.Logon
	Set oFolder = oNs.GetDefaultFolder(6)
	'oFolder.Display
	
	Set msg = oApp.CreateItem(0)
	Set Recpt = msg.Recipients.Add(sRecipients)
	msg.Subject = sSubject
	msg.body = sBody
	oRow.Cells(1, XL_OUTPUT_PARAMS).Value = "PARAM1=" & Now() & ",PARAM2=" & FormatDateTime(now(),2)
	msg.send
	
	
	sLog = sLog & "D: Message sent " & vbLf
	
	'WScript.Sleep(wait*1000)
	WaitTime((iWait/2)*1000) 
	oApp.quit
	WaitTime((iWait/2)*1000)
	
	
	Send = XL_DISPATCH_PASS
	oVBSFramework.oTraceLog.Exited(sRoutine)
End Function


'==============================================================================================
' Function/Sub: FindMailMail()
' Purpose:
' 	Search for message that match given "subject" param, resend previous message when sub match "resend" param
' 	Example For Fax Check test--->	
'		Zugestellt | param(TIMESENT,PARAM1) | TRUE | Unzustellbar | 60
' Parameters:
'	subject, time sent, resend bool, resend when found message with param4 matching
' Returns: delivery time from message body
'==============================================================================================
Function FindMail(oRow, sLog)
	Dim sSub, iTime, dTime, dDelivery, iDiff
	Dim oApp, oNs, oFolder, item1, item2, bResend, sResSub, iDelay
	
	Const sRoutine = "clsFaxCheck.FindMail"
	oVBSFramework.oTraceLog.Entered(sRoutine)
	FindMail = XL_DISPATCH_FAIL
	
	'parameter handling
	sSub = 		CStr(oRow.Cells(1, XL_PARM_001).Value)
	dTime = 	CDate(oRow.Cells(1, XL_PARM_002).Value)
	sResSub = 	CStr(oRow.Cells(1, XL_PARM_004).Value)
	iDelay = 	CDbl(oRow.Cells(1, XL_PARM_005).Value)
	
	Select Case UCase(CStr(oRow.Cells(1, XL_PARM_003).Value))
		Case "ON", "TRUE", "YES", "1"
			bResend = True						
		Case Else
			bResend = False
	End Select
	
	WaitTime((10)*1000)
	Set oApp=CreateObject("Outlook.Application")
	WaitTime(10000)
	Set oNs=oApp.GetNameSpace("MAPI")
	If oNs Is Nothing Then
		sLog = sLog & "E: Error launching Outlook Application" & vbLf
		SendMail = XL_DISPATCH_FAIL : oVBSFramework.oTraceLog.Exited(sRoutine) : Exit Function
	End If
	WaitTime((10)*1000)
	oNs.Logon
	Set oFolder=oNs.GetDefaultFolder(6)		'inbox
	WaitTime((10)*1000)
	
	iDiff = Timer
	Do While ((Timer-iDiff) < iDelay)
	  For each item1 in oFolder.Items
	  	'Zugestellt
		if UCase(Left(item1.subject,10)) = UCase(Left(sSub,10)) Then
			If dTime <= item1.ReceivedTime Then
				sLog = sLog & "D: Message found" & vbLf
				dDelivery = CDate(FormatDateTime(item1.ReceivedTime,4))
				'return +- 2 min
				oRow.Cells(1, XL_OUTPUT_PARAMS).Value = "PARAM1=" & Left(DateAdd("n",-2,dDelivery),5) & ",PARAM2=" & Left(DateAdd("n",2,dDelivery),5)
				FindMail = XL_DISPATCH_PASS
				Exit do
			End If
		ElseIf UCase(Left(item1.subject,10)) = UCase(Left(sResSub,10)) Then
			'Unzustellbar
			if bResend Then
				If dTime <= item1.ReceivedTime Then
					sLog = sLog & "D: Resending message" & vbLf
					bResend = False 'only one message sent
					dTime = Now()
					If XL_DISPATCH_FAIL = Send(oRow, sRecipients, sSubject, sBody, iWait, sLog) Then
						sLog = sLog & "E: Error resending message" & vbLf
						SendMail = XL_DISPATCH_FAIL : oVBSFramework.oTraceLog.Exited(sRoutine) : Exit Function
					End if
				End If
			end if
		End if
	  Next
	Loop
	
	oApp.Quit
	
	If FindMail <> XL_DISPATCH_PASS Then
		sLog = sLog & "E: Mail not found" & vbLf
	End If
	oVBSFramework.oTraceLog.Exited(sRoutine)
End Function
'==============================================================================================
' Function/Sub: TimeDif(iEndTime, iStartTime)
' Purpose: Avoid giving negative values when midnight passes
'
' Parameters: end time, start time
'
' Returns: elapsed time
'==============================================================================================
Function TimeDif(iEndTime, iStartTime)
	If iStartTime < iEndTime Then
		TimeDif = iEndTime - iStartTime
	Else
		TimeDif = iEndTime - iStartTime + 86400
	End If
End Function

'==============================================================================================
' Function/Sub: RegexReplace(sPattern, sSearchString, sReplaceString, bIgnoreCase)
' Purpose: RegexReplace matches the regular expression in the search-string, then it replaces that match with the replace-string, 
'				and the new string is returned
'==============================================================================================
Function RegexReplace(sPattern, sSearchString, sReplaceString, bIgnoreCase)
	Dim oRegEx
	
    Set oRegEx = New RegExp	
    
	oRegEx.Pattern = sPattern
	oRegEx.IgnoreCase = bIgnoreCase
	oRegEx.Global = True
	
	RegexReplace = oRegEx.Replace(sSearchString, sReplaceString)
	
	Set oRegEx = Nothing
End Function

'CONVERT FUNCTIONS:
'	example: ConvertLocalisedStringToDouble(sLocalizedString, "de-DE")
Private Function ConvertLocalisedStringToDouble(sLocalizedString, sLocale)
	Dim sCurrentLocaleID
	Dim dConvertedValue		
	sCurrentLocaleID = GetLocale()
'		SetLocale("de-DE")
	SetLocale(sLocale)
	dConvertedValue = CDbl(sLocalizedString)		 		
	SetLocale(sCurrentLocaleID) 
	ConvertLocalisedStringToDouble = dConvertedValue
End Function
	
'	example: ConvertLocalisedStringToDoubleDecPlaces(sLocalizedString, "de-DE", 3)
Private Function ConvertLocalisedStringToDoubleDecPlaces(sLocalizedString, sLocale, iDecimalPlaces)
	Dim sCurrentLocaleID
	Dim dConvertedValue		
	sCurrentLocaleID = GetLocale()
'		SetLocale("de-DE")
	SetLocale(sLocale)
	dConvertedValue = CDbl(sLocalizedString)		 		
	dConvertedValue = CDbl(FormatNumber(dConvertedValue, iDecimalPlaces))
	SetLocale(sCurrentLocaleID) 
	ConvertLocalisedStringToDoubleDecPlaces = dConvertedValue
End Function	
	
'	example: ConvertDoubleToLocalisedString(dValue, "de-DE")	
Private Function ConvertDoubleToLocalisedString(dValue, sLocale)
	Dim sCurrentLocaleID
	Dim dConvertedValue
	dValue = CDbl(dValue & "")
	sCurrentLocaleID = GetLocale()
'		SetLocale("de-DE")
	SetLocale(sLocale)
	dConvertedValue = CDbl(dValue)
	ConvertDoubleToLocalisedString = CStr(dConvertedValue)
	SetLocale(sCurrentLocaleID) 		
End Function
	
'	example: ConvertDoubleToLocalisedStringDecPlaces(dValue, "de-DE", 3)		
Private Function ConvertDoubleToLocalisedStringDecPlaces(dValue, sLocale, iDecimalPlaces)
	Dim sCurrentLocaleID
	Dim dConvertedValue
	Dim sTemp
	Dim i
	dValue = CDbl(dValue & "")
	sCurrentLocaleID = GetLocale()
	
'		SetLocale("de-DE")
	SetLocale(sLocale)
	dConvertedValue = CDbl(FormatNumber(dValue, iDecimalPlaces))	
	
	'force to add 'iDecimalPlaces' zeros to end of localised string, if dValue is whole number
	sTemp = CStr(dConvertedValue)
'		If Not(InStr(sTemp, ",") > 0) And iDecimalPlaces > 0 Then
'			sTemp = sTemp + ","
'			For i = 1 To iDecimalPlaces
'				sTemp = sTemp + "0"
'			Next
'		End If
	
	ConvertDoubleToLocalisedStringDecPlaces = sTemp
	SetLocale(sCurrentLocaleID) 
End Function

'==============================================================================================
'ROBODOC header blocks follow ...
'***if* Functions/ArrayFromIntervals,ArrayFromIntervals{Functions}
'  SYNOPSIS
'	Public Function ArrayFromIntervals(sIntervals) - create array of integers from string consists of single numbers and intervals
'  OVERVIEW
'	sIntervals - string with comma separated single numbers and intervals; i.e. "1-10, 11-11, 12, 14 -   20,  21 "
'*****
'==============================================================================================
Public Function ArrayFromIntervals(sIntervals)
	Dim arrIntervals, arrBounds
	Dim sInterval, sItems, sOutput
	Dim iLowBound, iUppBound, i
	
	sOutput = "!"	'special symbol to mark beginning
	arrIntervals = Split(sIntervals,",")
	For Each sInterval In arrIntervals
		arrBounds = Split(sInterval,"-")
		If UBound(arrBounds) = 0 Then	'single value
			sOutput = sOutput & "," & Trim(arrBounds(0))
		ElseIf UBound(arrBounds) = 1 Then	'interval
			iLowBound = Trim(arrBounds(0))
			iUppBound = Trim(arrBounds(1))
			For i = iLowBound To iUppBound
				sOutput = sOutput & "," & CStr(i)
			Next
		Else
			'Exit Function	'Wrong user input
		End If
	Next
	
	sOutput = Replace(sOutput,"!,", "")	'removing unwanted "!," string from the beginning
	ArrayFromIntervals = Split(sOutput,",")	
End Function
'==============================================================================================
' Sub: CheckForBackup(sLogFile)
' Purpose:
' 	Checking if sLogFile is not too big (more than 1GB).
'	If so, subroutine is backing up original file.
' 	
' Parameters:
'	sLogFile ...path to the log file
'==============================================================================================
Public Sub CheckForBackup(ByVal sLogFile)
	Dim iSize
	Dim oFS, oFile
	Dim iLog
	Const iGIGA = 1000000000
	
	sLogFile = Replace(sLogFile, """", "")
	
	Set oFS = CreateObject("Scripting.FileSystemObject")
	If oFS.FileExists(sLogFile) Then
		Set oFile = oFS.GetFile(sLogFile)
		If oFile.Size > iGIGA Then	'...it's time to back up things
			iLog = 1
			Do While oFS.FileExists(sLogFile & ".bckp" & iLog)
				iLog = iLog + 1
			Loop
			On Error Resume Next
				oFS.MoveFile oFile.Path, oFile.ParentFolder & "\" & oFile.Name & ".bckp" & iLog
				If Err.Number <> 0 Then
					'TODO: Need to handle "access rights" problem.
				End If
			On Error Goto 0
		End If
	'Else MsgBox sLogFile & " does not exist."
	End If
End Sub