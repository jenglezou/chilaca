'--This script demonstrates the window functions of JSSys3.dll.
' 
'--GetOpenWindowTitles 
'--GetActiveWindowTitle 
'--SetWindowActive 
'--CloseProgram 
 
'-- Other methods are demonstrated in the other script samples.

'-- this is a long script. The methods don't require so much code. They 're just
'-- written here to give full error reporting.

'************************************************************************
'------------------------------- START SCRIPT HERE:    -----------------------------------------

Dim Sys, v, s1, a1, i, r, s, sAct, boo
Set Sys = CreateObject("JSSys3.ops")

boo = False

s = "This script demonstrates window functions." & vbcrlf
s = s & "Before clicking OK:" & vbcrlf
s = s & "1 - Open Notepad." & vbcrlf
s = s & "2 - Open a folder window."

MsgBox s

'------------active window title -------------------------

MsgBox "1 - GetActiveWindowTitle..."

WScript.sleep 300 '--give a chance For MsgBox to lose focus so GetActiveWindowTitle doesn 't fail.

sAct = Sys.GetActiveWindowTitle()
 If sAct = "" Then
   MsgBox "Failed to get active window title. The message box in this script may be preventing an active window."
 Else
   MsgBox "Active window title is " & vbcrlf & sAct
 End If    

sAct = ""

'---------list of open windows ---------------------------------

MsgBox "2 - GetOpenWindowTitles..."

v = Sys.GetOpenWindowTitles(a1)
  If v > 0 Then
     s = "Open windows:" & vbcrlf & vbcrlf
    For i = 0 to UBound(a1)
      s = s & a1(i) & vbcrlf
    Next   
 Else
   s = "No open windows found."
 End If
 
 MsgBox s     

'---------- Set active window --------------------------

MsgBox "3 - SetWindowActive..."

v = Sys.GetOpenWindowTitles(a1)
  If v > 0 Then
    For i = 0 to UBound(a1)
       If InStr(1, a1(i), "Notepad", 1) <> 0 Then
          sAct = a1(i)
          Exit For
       End If   
    Next   
  End If  

     If sAct = "" Then
        MsgBox "Notepad was to be set active but it was not found."
     Else
       v = Sys.SetWindowActive(sAct, True)
       
         If v <> 0 Then
           Select Case v
             Case 1
               s = "No open windows found."
             Case 2
               s = "No window text found that matches string sent in SetActiveWindow."
              Case 3
               s = "Window is already active."
              Case 4
               s = "Attempt to Set active window failed."
              Case 5
               s = "Unknown error attempting to set active window."
            End Select
             MsgBox s
         End If
     End If
  
  

WScript.sleep 2000 '--wait a moment to display active.

'----------- close program --------------------

'--If Notepad string was found, use it to close Notepad:

 If sAct <> "" Then
   Sys.CloseProgram sAct, 2
  End If
  
  MsgBox "Notepad should have just closed. End of demo." 
       
Set Sys = Nothing