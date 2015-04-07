'--This script demonstrates several of the methods of JSSys3.dll.
'GetOpVersion 
'GetMemory 
'GetCurUser 
'GetCompName 
'GetDefaultProgram
'GetColorDepthBPP 
'GetDesktopPixels 
'GetScreenPixels 
'GetSystemColor     
'GetDriveInfo
'GetProcessList
'PlayWav

'-- Other methods are demonstrated in the other script samples.

'--NOTE: gong.wav needs to be on Desktop For this script (or rewrite the script).--------
'************************************************************************
'------------------------------- START SCRIPT HERE:    -----------------------------------------

Dim Sys, v, SH, t, u, s, s1, w, h, a1, i, num
Set Sys = CreateObject("JSSys3.ops")
Set SH = CreateObject("WScript.Shell")

'--------------------------------Get WINDOWS VERSION.
MsgBox "1 - GetOpVersion..."

v = Sys.GetOpVersion(a1)
 If v = 0 Then
  s = "Platform: " & a1(0) & vbcrlf
  s = s & "Major Version: "  & a1(1) & vbcrlf
   s = s & "Minor Version: "  & a1(2) & vbcrlf
   s = s & "Build Number: "  & a1(3) & vbcrlf
   s = s & "Extra info. or service pack: "  & a1(4) 
 Else
   s = "Call to GetOpVersion failed."
 End If
MsgBox s

'------------------------------Get PHYS. MEMORY AND % IN USE.
MsgBox "2 - GetMemory..."

v = Sys.GetMemory(t, u)
 If v = 0 Then
    s = "Total RAM: " & t & vbcrlf & "Current RAM in use: " & u
 Else
    s = "Call to GetMemory failed."   
 End If   
 
MsgBox s

'-----------------------------------Get CURRENT USER NAME.
MsgBox "3 - GetCurUser..."
v = Sys.GetCurUser
MsgBox "The current user is: " & v

'------------------------------------Get COMPUTER NAME.
MsgBox "4 - GetCompName..."
v = Sys.GetCompName
MsgBox "The computer name is: " & v

'-----------------------------------Get DEFAULT PROGRAM.
MsgBox "5 - GetDefaultProgram..."
v = Sys.GetDefaultProgram("bmp")
MsgBox "The default program for file extension BMP is" & vbcrlf & v

'-----------------------------------------------------------Get DISPLAY DEPTH.
MsgBox "6 - GetColorDepthBPP..."

v = Sys.GetColorDepthBPP
 If v <> 0 Then
    MsgBox "The current display depth is " & v & " bits per pixel."
 Else
    MsgBox " Call to GetColorDepthBPP failed."
  End If
'--------------------------------------------------------Get WORKING AREA SIZE,WITHOUT TASKBAR.

MsgBox "7 - GetDesktopPixels..."
v = Sys.GetDesktopPixels(w, h)
If v = 0 Then
  s = "The Desktop working area is " & w & " pixels wide and " & h & " pixels high."
Else
 s = "Call to GetDesktopPixels failed."
End If
 MsgBox s
 
'---------------------------------------------------Get SCREEN SIZE.

MsgBox "8 - GetScreenPixels..."
v = Sys.GetscreenPixels(w, h)
If v = 0 Then
  s = "The screen area is " & w & " pixels wide and " & h & " pixels high."
Else
 s = "Call to GetScreenPixels failed."
End If
 MsgBox s

'-----------------------------------------------Get DRIVE INFO For ALL FIXED DRIVES.
MsgBox "9 - GetDriveInfo..."

v = Sys.GetDriveInfo
  If v = "" Then
    s = "Call to GetDriveInfo failed."
  Else
     a1 = split(v, ",")
     s = "DRIVE-TOTAL MB-FREE" & vbcrlf & vbcrlf
     For i = 0 to ubound(a1)
       s = s & a1(i) & vbcrlf
     Next
  End If   
 MsgBox s
 
'-------------------------------------------- Get system colors.
MsgBox "10 - GetSystemColor..."
  
  v = Sys.GetSystemColor("Buttons")
     If len(v) = 6 Then
       s = "Hex code for color of window frames and buttons is " & v
     Else
       s = "Call to GetSystemColor failed."
     End If
  MsgBox s
  
'------------------------------------------ Get list of processes running on 95/98/ME/2000/XP
MsgBox "11 - GetProcessList"
 v = Sys.GetProcessList(s, num)
   Select Case v
     Case 1
       s = "Unable to confirm Windows version."
     Case 2
       s = "This is not a Win9x/2000/XP system. GetProcessList cannot be used."
     Case 3
       s = "Call to GetProcessList failed."
     Case 0
       a1 = split(s, ",")
       For i = 0 to UBound(a1)
         s1 = s1 & a1(i) & vbcrlf
       Next    
        s = "There are " & num & " processes running:" & vbcrlf & "(This may not all fit into a message box.)" & vbcrlf & vbcrlf
        s = s & s1
    End Select
      MsgBox s    
      
'----------------------------------------------PLAY SOUND FILE (WAV)
MsgBox "Final demo in this script: PlayWav..."
'-----------make sure the path is correct For this demo.
s = SH.SpecialFolders("Desktop")
Sys.PlayWav s & "\gong.wav"

Set SH = Nothing
Set Sys = Nothing


