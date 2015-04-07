Option Explicit

Const HKEY_CURRENT_USER = &H80000001
Const HKEY_LOCAL_MACHINE = &H80000002

Dim sComputer, sINIFile, sProgramFilesPath
Dim sKeyPath, sKeyName, skeyValue, dwKeyValue
Dim oReg, oQTP, oShell, oFS
Dim sLog

sLog = now() & vbNewLine
sComputer = "."
 
Set oReg=GetObject("winmgmts:{impersonationLevel=impersonate}!\\" & _
    sComputer & "\root\default:StdRegProv")

' ************************************   JET 4.0
sKeyPath = "SOFTWARE\Microsoft\Jet\4.0\Engines\Text"
sKeyName = "Format"
oReg.GetStringValue HKEY_LOCAL_MACHINE, sKeyPath, sKeyName, skeyValue
sLog = sLog & vbNewLine & "BEFORE: " & sKeyPath & "\" & sKeyName & ":" & skeyValue & vbNewLine
If skeyValue <> "Delimited(,)" Then 
	oReg.SetStringValue HKEY_LOCAL_MACHINE, sKeyPath, sKeyName, "Delimited(,)"
	oReg.GetStringValue HKEY_LOCAL_MACHINE, sKeyPath, sKeyName, skeyValue
	sLog = sLog & "After: " & sKeyPath & "\" & sKeyName & ":" & skeyValue & vbNewLine
Else
	sLog = sLog & "AFTER: No change required." & vbNewLine
End If

' ************************************   QTP mic.ini update for patch 0626
Set oShell = CreateObject("WScript.Shell")
sProgramFilesPath = oShell.ExpandEnvironmentStrings("%ProgramFiles%")
Set oShell = Nothing

sINIFile = sProgramFilesPath & "\HP\QuickTest Professional\bin\mic.ini"
Set oFS = CreateObject("Scripting.Filesystemobject")

If oFS.FileExists(sINIFile) Then

	If ReadIni(sINIFile, "RemoteAgent", "HandleLongTrans") <> "1" Then
		'insert value
		sLog = sLog & vbNewLine & "mic.ini - setting up HandleLongTrans=1 into [RemoteAgent] section." & vbNewLine
		WriteIni sINIFile, "RemoteAgent", "HandleLongTrans", "1"
	Else
		sLog = sLog & vbNewLine & "mic.ini - HandleLongTrans=1 already set into [RemoteAgent] section." & vbNewLine
	End If

'NOTE: Extern object not supported by VBScript (only by QTP)
'	Extern.Declare micInteger, "GetPrivateProfileStringA", "kernel32.dll", "GetPrivateProfileStringA", micString, micString, micString, micString+micByRef, micInteger, micString
'	Extern.Declare micLong, "WritePrivateProfileStringA", "kernel32.dll", "WritePrivateProfileStringA", micString, micString, micString,  micString
'	If CInt(Extern.GetPrivateProfileStringA("RemoteAgent", "HandleLongTrans", "N/A", sKeyValue, 255, sINIFile)) = 1 And sKeyValue = "1" Then
'		sLog = sLog & vbNewLine & "mic.ini - HandleLongTrans=1 already set into [RemoteAgent] section." & vbNewLine
'	Else
'		sLog = sLog & vbNewLine & "mic.ini - setting up HandleLongTrans=1 into [RemoteAgent] section." & vbNewLine
'		Extern.WritePrivateProfileStringA "RemoteAgent", "HandleLongTrans", "1", sINIFile
'	End If
Else
	sLog = sLog & vbNewLine & "mic.ini - File not found: " & sINIFile & vbNewLine
End If

Set oFS = Nothing

' ************************************   QTP
sKeyPath = "Software\Mercury Interactive\QuickTest Professional\MicTest"
sKeyName = "AllowTDConnect"
oReg.GetDWORDValue HKEY_CURRENT_USER, sKeyPath, sKeyName, dwKeyValue

'Start QTP if it was never started before under current user to correctly create registry structure
If IsNull(dwKeyValue) Then
	sLog = sLog & vbNewLine & "The registry key does not exist: " & sKeyPath & "\" & sKeyName & vbNewLine
	Set oQTP = CreateObject("QuickTest.Application")
	oQTP.Launch
	oQTP.Visible = True
	oQTP.Quit
	Set oQTP = Nothing
End If

oReg.GetDWORDValue HKEY_CURRENT_USER, sKeyPath, sKeyName, dwKeyValue
sLog = sLog & vbNewLine & "BEFORE: " & sKeyPath & "\" & sKeyName & ":" & dwKeyValue & vbNewLine
If dwKeyValue = 0 Then 
	oReg.SetDWORDValue HKEY_CURRENT_USER, sKeyPath, sKeyName, 1
	oReg.GetDWORDValue HKEY_CURRENT_USER, sKeyPath, sKeyName, dwKeyValue
	sLog = sLog & "After: " & sKeyPath & "\" & sKeyName & ":" & dwKeyValue & vbNewLine
Else
	sLog = sLog & "AFTER: No change required." & vbNewLine
End If

sKeyName = "RunMode"
oReg.GetDWORDValue HKEY_CURRENT_USER, sKeyPath, sKeyName, dwKeyValue
sLog = sLog & vbNewLine & "BEFORE: " & sKeyPath & "\" & sKeyName & ":" & dwKeyValue & vbNewLine
If dwKeyValue = 0 Then 
	oReg.SetDWORDValue HKEY_CURRENT_USER, sKeyPath, sKeyName, 1
	oReg.GetDWORDValue HKEY_CURRENT_USER, sKeyPath, sKeyName, dwKeyValue
	sLog = sLog & "AFTER: " & sKeyPath & "\" & sKeyName & ":" & dwKeyValue & vbNewLine
Else
	sLog = sLog & "AFTER: No change required." & vbNewLine
End If

sKeyName = "LaunchReport"
oReg.GetDWORDValue HKEY_CURRENT_USER, sKeyPath, sKeyName, dwKeyValue
sLog = sLog & vbNewLine & "BEFORE: " & sKeyPath & "\" & sKeyName & ":" & dwKeyValue & vbNewLine
If dwKeyValue = 1 Then 
	oReg.SetDWORDValue HKEY_CURRENT_USER, sKeyPath, sKeyName, 0
	oReg.GetDWORDValue HKEY_CURRENT_USER, sKeyPath, sKeyName, dwKeyValue
	sLog = sLog & "AFTER: " & sKeyPath & "\" & sKeyName & ":" & dwKeyValue & vbNewLine
Else
	sLog = sLog & "AFTER: No change required." & vbNewLine
End If


' ****************************** Desktop Wallpaper
sKeyPath = "Control Panel\Desktop"
sKeyName = "Wallpaper"
oReg.GetStringValue HKEY_CURRENT_USER, sKeyPath, sKeyName, sKeyValue
sLog = sLog & vbNewLine & "BEFORE: " & sKeyPath & "\" & sKeyName & ":" & sKeyValue & vbNewLine
If sKeyValue <> "" Then 
	oReg.SetStringValue HKEY_CURRENT_USER, sKeyPath, sKeyName, ""
	oReg.GetStringValue HKEY_CURRENT_USER, sKeyPath, sKeyName, sKeyValue
	sLog = sLog & "AFTER: " & sKeyPath & "\" & sKeyName & ":" & sKeyValue & vbNewLine
	sLog = sLog & "IMPORTANT: The desktop wallpaper change requires a logoff/logon." & vbNewLine
Else
	sLog = sLog & "AFTER: No change required." & vbNewLine
End If


' ******************************  Screen Saver
sKeyPath = "Control Panel\Desktop"
sKeyName = "ScreenSaveActive"
oReg.GetStringValue HKEY_CURRENT_USER, sKeyPath, sKeyName, sKeyValue
sLog = sLog & vbNewLine & "BEFORE: " & sKeyPath & "\" & sKeyName & ":" & sKeyValue & vbNewLine
If sKeyValue <> "0" Then 
	oReg.SetStringValue HKEY_CURRENT_USER, sKeyPath, sKeyName, "0"
	oReg.GetStringValue HKEY_CURRENT_USER, sKeyPath, sKeyName, sKeyValue
	sLog = sLog & "AFTER: " & sKeyPath & "\" & sKeyName & ":" & sKeyValue & vbNewLine

	sKeyName = "SCRNSAVE.EXE"
	oReg.GetStringValue HKEY_CURRENT_USER, sKeyPath, sKeyName, sKeyValue
	sLog = sLog & vbNewLine & "BEFORE: " & sKeyPath & "\" & sKeyName & ":" & sKeyValue & vbNewLine
	oReg.DeleteValue HKEY_CURRENT_USER, sKeyPath, sKeyName
	oReg.GetStringValue HKEY_CURRENT_USER, sKeyPath, sKeyName, sKeyValue
	sLog = sLog & "AFTER: " & sKeyPath & "\" & sKeyName & ":" & sKeyValue & vbNewLine
Else
	sLog = sLog & "AFTER: No change required." & vbNewLine
End If

' ****************************** Load Unified Functional Testing License
sKeyPath = "Software\Mercury Interactive\License Manager\UFT"
sKeyName = "Need" 
oReg.GetDWORDValue HKEY_CURRENT_USER, sKeyPath, sKeyName, dwKeyValue
sLog = sLog & vbNewLine & "BEFORE: " & sKeyPath & "\" & sKeyName & ":" & dwKeyValue & vbNewLine
If dwKeyValue = 0 Then 
	oReg.SetDWORDValue HKEY_CURRENT_USER, sKeyPath, sKeyName, 1
	oReg.GetDWORDValue HKEY_CURRENT_USER, sKeyPath, sKeyName, dwKeyValue
	sLog = sLog & "AFTER: " & sKeyPath & "\" & sKeyName & ":" & dwKeyValue & vbNewLine
Else
	sLog = sLog & "AFTER: No change required." & vbNewLine
End If

Function ReadIni(myFilePath, mySection, myKey)
    Const ForReading   = 1
    Const ForWriting   = 2
    Const ForAppending = 8

    Dim intEqualPos
    Dim objFSO, objIniFile
    Dim strFilePath, strKey, strLeftString, strLine, strSection

    Set objFSO = CreateObject( "Scripting.FileSystemObject" )

    ReadIni     = ""
    strFilePath = Trim( myFilePath )
    strSection  = Trim( mySection )
    strKey      = Trim( myKey )

    If objFSO.FileExists( strFilePath ) Then
        Set objIniFile = objFSO.OpenTextFile( strFilePath, ForReading, False )
        Do While objIniFile.AtEndOfStream = False
            strLine = Trim( objIniFile.ReadLine )

            ' Check if section is found in the current line
            If LCase( strLine ) = "[" & LCase( strSection ) & "]" Then
                strLine = Trim( objIniFile.ReadLine )

                ' Parse lines until the next section is reached
                Do While Left( strLine, 1 ) <> "["
                    ' Find position of equal sign in the line
                    intEqualPos = InStr( 1, strLine, "=", 1 )
                    If intEqualPos > 0 Then
                        strLeftString = Trim( Left( strLine, intEqualPos - 1 ) )
                        ' Check if item is found in the current line
                        If LCase( strLeftString ) = LCase( strKey ) Then
                            ReadIni = Trim( Mid( strLine, intEqualPos + 1 ) )
                            ' In case the item exists but value is blank
                            If ReadIni = "" Then
                                ReadIni = " "
                            End If
                            ' Abort loop when item is found
                            Exit Do
                        End If
                    End If

                    ' Abort if the end of the INI file is reached
                    If objIniFile.AtEndOfStream Then Exit Do

                    ' Continue with next line
                    strLine = Trim( objIniFile.ReadLine )
                Loop
            Exit Do
            End If
        Loop
        objIniFile.Close
    Else
        WScript.Echo strFilePath & " doesn't exists. Exiting..."
        Wscript.Quit 1
    End If
End Function

Sub WriteIni(myFilePath, mySection, myKey, myValue)
    Const ForReading   = 1
    Const ForWriting   = 2
    Const ForAppending = 8

    Dim blnInSection, blnKeyExists, blnSectionExists, blnWritten
    Dim intEqualPos
    Dim objFSO, objNewIni, objOrgIni, wshShell
    Dim strFilePath, strFolderPath, strKey, strLeftString
    Dim strLine, strSection, strTempDir, strTempFile, strValue

    strFilePath = Trim( myFilePath )
    strSection  = Trim( mySection )
    strKey      = Trim( myKey )
    strValue    = Trim( myValue )

    Set objFSO   = CreateObject( "Scripting.FileSystemObject" )
    Set wshShell = CreateObject( "WScript.Shell" )

    strTempDir  = wshShell.ExpandEnvironmentStrings( "%TEMP%" )
    strTempFile = objFSO.BuildPath( strTempDir, objFSO.GetTempName )

    Set objOrgIni = objFSO.OpenTextFile( strFilePath, ForReading, True )
    Set objNewIni = objFSO.CreateTextFile( strTempFile, False, False )

    blnInSection     = False
    blnSectionExists = False
    ' Check if the specified key already exists
    blnKeyExists     = ( ReadIni( strFilePath, strSection, strKey ) <> "" )
    blnWritten       = False

    ' Check if path to INI file exists, quit if not
    strFolderPath = Mid( strFilePath, 1, InStrRev( strFilePath, "\" ) )
    If Not objFSO.FolderExists ( strFolderPath ) Then
        WScript.Echo "Error: WriteIni failed, folder path (" _
                   & strFolderPath & ") to ini file " _
                   & strFilePath & " not found!"
        Set objOrgIni = Nothing
        Set objNewIni = Nothing
        Set objFSO    = Nothing
        WScript.Quit 1
    End If

    While objOrgIni.AtEndOfStream = False
        strLine = Trim( objOrgIni.ReadLine )
        If blnWritten = False Then
            If LCase( strLine ) = "[" & LCase( strSection ) & "]" Then
                blnSectionExists = True
                blnInSection = True
            ElseIf InStr( strLine, "[" ) = 1 Then
                blnInSection = False
            End If
        End If

        If blnInSection Then
            If blnKeyExists Then
                intEqualPos = InStr( 1, strLine, "=", vbTextCompare )
                If intEqualPos > 0 Then
                    strLeftString = Trim( Left( strLine, intEqualPos - 1 ) )
                    If LCase( strLeftString ) = LCase( strKey ) Then
                        ' Only write the key if the value isn't empty
                        ' Modification by Johan Pol
                        If strValue <> "<DELETE_THIS_VALUE>" Then
                            objNewIni.WriteLine strKey & "=" & strValue
                        End If
                        blnWritten   = True
                        blnInSection = False
                    End If
                End If
                If Not blnWritten Then
                    objNewIni.WriteLine strLine
                End If
            Else
                objNewIni.WriteLine strLine
                    ' Only write the key if the value isn't empty
                    ' Modification by Johan Pol
                    If strValue <> "<DELETE_THIS_VALUE>" Then
                        objNewIni.WriteLine strKey & "=" & strValue
                    End If
                blnWritten   = True
                blnInSection = False
            End If
        Else
            objNewIni.WriteLine strLine
        End If
    Wend

    If blnSectionExists = False Then ' section doesn't exist
        objNewIni.WriteLine
        objNewIni.WriteLine "[" & strSection & "]"
            ' Only write the key if the value isn't empty
            ' Modification by Johan Pol
            If strValue <> "<DELETE_THIS_VALUE>" Then
                objNewIni.WriteLine strKey & "=" & strValue
            End If
    End If

    objOrgIni.Close
    objNewIni.Close

    ' Delete old INI file
    objFSO.DeleteFile strFilePath, True
    ' Rename new INI file
    objFSO.MoveFile strTempFile, strFilePath

    Set objOrgIni = Nothing
    Set objNewIni = Nothing
    Set objFSO    = Nothing
    Set wshShell  = Nothing
End Sub

'WScript.Echo sLog
