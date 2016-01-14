option explicit

'==============================================================================================
'Check if a DLL/EXE is already loaded. 
Private Function IsServerLoaded(sObject)
	Dim oObject
	Dim bRetVal
	
	bRetVal = False
	'Check if the DLL/EXE is already loaded before trying to load it
	On error resume next
	set oObject = createobject(sObject)
	On Error GoTo 0	'Turn off on error handling
	If isobject(oObject) then
		'MsgBox  "DLL containing " & sObject & " is loaded."
		bRetVal = True
	else
		'MsgBox  "DLL containing " & sObject & " is not loaded."
	End if

	IsServerLoaded = bRetVal
End Function

'=================================================================================================
'Load an EXE. Unload it if required. 
Private Sub LoadEXE(sEXEFile, sObject2Check, bUnloadFirst)
	
	If bUnloadFirst = True then ExecuteCommand sEXEFile & "/unregserver", True, True, sVBSFrameworkDir & "\temp\ExecuteCommand.tmp" 'Unload
		
	If Not IsServerLoaded(sObject2Check) then 
		ExecuteCommand sEXEFile & " /regserver", True, True, sVBSFrameworkDir & "\temp\ExecuteCommand.tmp"
	End If
End Sub

'=================================================================================================
'Load an DLL. Unload it if required. 
Private Sub LoadDLL(sDLLFile, sObject2Check, bUnloadFirst)
	
	If bUnloadFirst = True then ExecuteCommand "regsvr32 /s /u " & sDLLFile, True, True, sVBSFrameworkDir & "\temp\ExecuteCommand.tmp"	'Unload
	
	If Not IsServerLoaded(sObject2Check) then 
		ExecuteCommand "regsvr32 /s " & sDLLFile, True, True, sVBSFrameworkDir & "\temp\ExecuteCommand.tmp"
	End If
End Sub
'=================================================================================================

'=================================================================================================
'Load an .NET assembly DLL. Unload it if required. 
Private Sub LoadNETDLL(sDLLFile, sTLBFile, sObject2Check, bUnloadFirst)
	
	If bUnloadFirst = True Then ExecuteCommand """c:\windows\microsoft.net\framework\v2.0.50727\regasm.exe"" " & sDLLFile & "/unregister", True, True, sVBSFrameworkDir & "\temp\ExecuteCommand.tmp"	'Unload
'STOP	
	If Not IsServerLoaded(sObject2Check) then 
		ExecuteCommand """c:\windows\microsoft.net\framework\v2.0.50727\regasm.exe"" " & sDLLFile & " /tlb:" & sTLBFile & " /codebase", True, True, sVBSFrameworkDir & "\temp\ExecuteCommand.tmp"
	End If
End Sub
'=================================================================================================

Dim sLibraryFolder	'Folder where the DLL/EXE's are.
Dim sEXEFile		'Set this to the exe file to be loaded.
Dim sDLLFile		'Set this to the dll file to be loaded.
Dim sTLBFile
Dim bUnload			'True means - force unload before loading.

'Set the library path
If  len(sVBSFrameworkDir) > 0 Then
	sLibraryFolder = sVBSFrameworkDir & "\Libraries\"
else
	sLibraryFolder = "c:\chilaca\Libraries\"
End If

bUnload = False

'sEXEFile = """" & sLibraryFolder & "JEUtilities.exe"""
'LoadEXE sEXEFile, "JEUtilities.itemlist", bUnload

'sDLLFILE = """" & sLibraryFolder & "JSSys3.dll"""
'LoadDLL sDLLFile, "JSSys3.ops", bUnload

'sDLLFILE = """" & sLibraryFolder & "TAUtility.dll"""
'sTLBFile = """" & sLibraryFolder & "TAUtility.tlb"""
'LoadNETDLL sDLLFile,sTLBFile,"TAUtility", bUnload

'sDLLFile = """" & sLibraryFolder & "FileDiffNET.DLL"""
'sTLBFile = """" & sLibraryFolder & "FileDiffNET.TLB"""
'LoadNETDLL sDLLFile, sTLBFile, "FileDiffNET.Compare", bUnload

'sDLLFile = """" & sLibraryFolder & "Selenium.DLL"""
'sTLBFile = """" & sLibraryFolder & "Selenium.TLB"""
'LoadNETDLL sDLLFile, sTLBFile, "Selenium.ChromeDriver", bUnload