'--This script demonstrates several of the clipboard methods of JSSys3.dll.
'SendTextCB
'GetTextCB

'-- Other methods are demonstrated in the other script samples.
'************************************************************************
'------------------------------- START SCRIPT HERE:    -----------------------------------------

Dim Sys, v, s, s1
Set Sys = CreateObject("JSSys3.ops")

s = InputBox("Enter text to send to Clipboard", "Clipboard Demo")
  If s = "" Then WScript.Quit
  
v = Sys.SendTextCB(s)
  If v = 0 Then
    MsgBox "Text is now on Clipboard."
  Else
    MsgBox "Error sending text to Clipboard."
  End If
  
MsgBox "Next, put some text on Clipboard and then click OK."
v = Sys.GetTextCB(s)
   If v = 1 Then
     s = "Clipboard data is not in text format."
   ElseIf v = 2 Then
     s = "Error calling GetTextCB"
   End If
 MsgBox s
   
  
Set Sys = Nothing