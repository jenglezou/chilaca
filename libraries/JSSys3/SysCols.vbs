
'-- demo to change system colors.

Dim Sys, s, TitC, TitTC, NewTitc, NewTitTC, v
Set Sys = CreateObject("JSSys3.Ops")

s = "This demo will change system colors. For convenience" & vbcrlf
s = s & "it will first get the current color in case you want to" & vbcrlf
s = s & "change it back."

MsgBox s

  TitC = Sys.GetSystemColor("TitleBar")
  TitTC = Sys.GetSystemColor("TitleBarText")
   
 If Len(TitC) = 6 and Len(TitTC) = 6 Then
    MsgBox "Title bar color is " & TitC & ". Title bar text color is " & TitTC
 Else
    MsgBox "Error getting colors."
    WScript.Quit
 End If

  NewTitc = InputBox("Enter 6-character hex code for new Title Bar color.", "System Colors")
     If Len(NewTitc) <> 6 Then
       MsgBox "Invalid. Script will quit."
       WScript.Quit
     End If
     
   NewTitTC = InputBox("Enter 6-character hex code for new Title Bar text color.", "System Colors")
     If Len(NewTitTC) <> 6 Then
       MsgBox "Invalid. Script will quit."
       WScript.Quit
     End If
  '------------------------Set title bar -------------------
     
 v = Sys.SetSystemColor("TitleBar", NewTitc)
     Select Case v 
        Case 0 
           s = "New title bar color set."
       Case 1
           s = "Invalid parameter for item."
       Case 2
           s = "Invalid hex code."
       Case 3
           s = "Failed to make color change."
       Case 4
           s = "Unable to make color change permanent. There may be an issue with permissions."
       Case 5
           s = "Unknown error attempting to set Title Bar color."
     End Select

msgbox "Title Bar Color:" & vbcrlf & s
                    
'-------------------Set T.B. text --------------------------

v = Sys.SetSystemColor("TitleBarText", NewTitTC)
     Select Case v 
        Case 0 
           s = "New title bar text color set."
       Case 1
           s = "Invalid parameter for item."
       Case 2
           s = "Invalid hex code."
       Case 3
           s = "Failed to make color change."
       Case 4
           s = "Unable to make color change permanent. There may be an issue with permissions."
       Case 5
           s = "Unknown error attempting to set Title Bar Text color."
     End Select

msgbox "Title Bar text:" & vbcrlf & s
     
MsgBox "If color change was successful you may want to change the settings back. Click OK when ready to proceed."

v = MsgBox("Do you want to reset the colors changed?", 36)
  If v = 7 Then
    Set Sys = Nothing
    WScript.quit
 End If
 
 v = Sys.SetSystemColor("TitleBar", TitC)  
 v = Sys.SetSystemColor("TitleBarText", TitTC)
 
MsgBox "Done."
Set Sys = Nothing
                    
     