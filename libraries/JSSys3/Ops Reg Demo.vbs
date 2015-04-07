'--This script demonstrates the timer and registry methods of JSSys3.dll.
 '--it will write values to HKCU\Software\TEST1 key. That key can be
 '--deleted after test.
 
'TimerStart 
'TimerStop 
'RegGetValue 
'RegWriteValue 
'RegListKeys 
'RegListVals 
'RegListStringData 
'RegDeleteSubkey

'-- Other methods are demonstrated in the other script samples.
'************************************************************************
'------------------------------- START SCRIPT HERE:    -----------------------------------------

Dim Sys, v, s1, a1, i, r, s2
Set Sys = CreateObject("JSSys3.ops")

'-----------timer.
MsgBox "Timer demo..."

Sys.TimerStart
WScript.sleep 400
v = Sys.TimerStop
  If v = 0 Then
    MsgBox "Timer Function failed."
  Else
    MsgBox "Time elapsed: " & v & " milliseconds."
  End If
   
'---------------------------------------------Get a LIST OF REG SUBKEYS (as array)
r = MsgBox("Registry demo. This demo will need to create a key HKCU\Software\TEST1. That key can be deleted after demo. Proceed?", 36)
  If r = 7 Then WScript.quit
  
MsgBox "RegListKeys..."

v = Sys.RegListKeys("HKCU", "Control Panel", a1)
 
 Select Case v
    Case 4
      s2 = "Error with RegListKeys."
    Case 3
      s2 = "Error. HKey parameter invalid."
    Case 2
      s2 = "Error. Failed to open key. Path may not exist."
    Case 1
      s2 = "There are no subkeys in that key."
    Case 0
      s2 = "Subkeys in HKCU\Control Panel key:" & vbcrlf & vbcrlf
        For i = 0 to ubound(a1)
          s2 = s2 & a1(i) & vbcrlf
        Next
   End Select
  
 MsgBox s2           
 MsgBox "RegListVals..."
'-------------------------------------------------Get a LIST OF REG KEY VALUES WITH TYPE (as array)
v = Sys.RegListVals("HKLM", "Software\Microsoft\Windows\CurrentVersion", a1)

    Select Case v
    Case 4
      s2 = "Error with RegListVals."
    Case 3
      s2 = "Error. HKey parameter invalid."
    Case 2
      s2 = "Error. Failed to open key. Path may not exist."
    Case 1
      s2 = "There are no values in that key."
    Case 0
       s2 = "Values in HKLM\Software\Microsoft\Windows\CurrentVersion" & vbcrlf & "with type indicator appended:" & vbcrlf & vbcrlf
        For i = 0 to ubound(a1)
          s2 = s2 & a1(i) & vbcrlf
        Next
   End Select
  
 MsgBox s2  
 MsgBox "RegListStringData..."      
 
'------------------------------------------------  Get a list of string values in key.

 v = Sys.RegListStringData("HKLM", "Software\Microsoft\Windows\CurrentVersion", a1)
    Select Case v
    Case 4
      s2 = "Error with RegListStringData."
    Case 3
      s2 = "Error. HKey parameter invalid."
    Case 2
      s2 = "Error. Failed to open key. Path may not exist."
    Case 1
      s2 = "There are no string values in that key."
    Case 0
       s2 = "String values and data in HKLM\Software\Microsoft\Windows\CurrentVersion" & vbcrlf & vbcrlf
          For i = 0 to ubound(a1)
            s2 = s2 & a1(i) & vbcrlf
          Next
       s2 = Replace(s2, "^", "  ")
   End Select
  
 MsgBox s2        
 
MsgBox "RegWriteValue..."
 '------------------------------------------- WRITE REGISTRY VALUE -----------------------
 MsgBox "First a REG_EXPAND_SZ value of %WinDir% will be written and then read back.", 64
 
 s1 = "%WinDir%"
 v = Sys.RegWriteValue("HKCU", "Software\TEST1", "TestValueSX", s1, "X")
    If (v <> 0) Then
        MsgBox "Failed to write REG_EXPAND_SZ value. Error: " & CStr(v)
    Else  
        v = Sys.RegGetValue("HKCU", "Software\TEST1", "TestValueSX", s1, "X")
          If v = 0 Then
              MsgBox "Expand string value: " & s1
          Else
              MsgBox "Error reading REG_EXPAND_SZ value just written. Error: " & v
          End If  
   End If       
 '--writes an array of binary data using hex format:
   MsgBox "Next - writing a binary value...", 64
   
 a1 = array("55", "FF", "FF", "00", "22", "56", "CC", "A0", "1D")
 v = Sys.RegWriteValue("HKCU", "Software\TEST1", "TestValueB", a1, "BH")
    Select Case v
      Case 0
        s2 = "Data written successfully."
      Case 1
        s2 = "Failed to open/create key."
      Case 2
        s2 = "Failed to set value."
      Case 3
        s2 = "Invalid binary data."
      Case 4
        s2 = "vKey value Not valid."
      Case 5
        s2 = "vRegPath value missing."
      Case 6
        s2 = "Attempting to set default value with data other than string."
      Case 7
        s2 = "Invalid DWord data."
      Case 8
        s2 = "Binary data Not in array."
      Case 9
        s2 = "Invalid type parameter."  
      Case 10
        s2 = "Unknown error."
   End Select         
     MsgBox s2
   
 If v = 0 Then  
   MsgBox "Final demo. RegGetValue will read the value just written."
 Else
   MsgBox "The script was going to read the value just written but there was an error with RegWriteValue."
   WScript.Quit
 End If
 
'-------------------- --------------------  READ A REGISTRY VALUE:    

v = Sys.RegGetValue("HKCU", "Software\TEST1", "TestValueB", a1, s1)
  Select Case v
    Case 6
      s2 = "Unknown error in attempting to get value."
    Case 5
      s2 = "Unknown value type."
     Case 4
      s2 = "HKEY parameter is not valid."
     Case 3
       s2 = "Value does not exist."
     Case 2
      s2 = "Failed to open key. Key path may be wrong."
     Case 1
      s2 = "Value contains no data."
     Case 0
        Select Case s1
          Case "S"
             s2 = "String data returned is:" & vbcrlf & a1
          Case "D"
             s2 = "DWord data returned is:" & vbcrlf & a1
          Case "B"
             s2 = "Binary data returned is:" & vbcrlf
           For i = 0 to UBound(a1)
                s2 = s2 & a1(i) & ", "
          Next      
               s2 = left(s2, (len(s2) - 2))     
       End Select
  End Select
       
     MsgBox s2
 r =  MsgBox("End of Registry demo. Delete TEST1 key just written?", 36)
    If r = 6 Then
      r = Sys.RegDeleteSubkey("HKCU", "Software", "TEST1")
        MsgBox "return from RegDeleteSubkey: " & r
   end if     
Set Sys = Nothing