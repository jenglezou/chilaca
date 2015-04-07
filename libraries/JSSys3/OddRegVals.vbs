'--demonstrates multi-string and long values with RegGetValue and RegWriteValue

Dim Sys, a, a1, r, s, i, vType
Set Sys = CreateObject("JSSys3.Ops")

a = array("value1", "value2", "value3", "value4")
r = Sys.RegWriteValue("HKCU", "Software\TestM", "MVal", a, "M")
  MsgBox "Return from RegWriteValue: " & r

r = Sys.RegGetValue("HKCU", "Software\TestM", "MVal", a1, vType)

 If r = 0 Then
       For i = 0 to UBound(a1)
         s = s & a1(i) & vbcrlf
       Next
         s = "Values:" & vbcrlf & s & vbcrlf & "Type: " & vtype
 Else
    s = "woops"
 End If
 
MsgBox "Return from RegGetValue: " & r & vbcrlf & vbcrlf & s
   

r = Sys.RegWriteValue("HKCU", "Software\TestM", "DVal1", 54699, "D")
 MsgBox "Return from RegWriteValue for DWord: " & r

r = Sys.RegWriteValue("HKCU", "Software\TestM", "DVal2", "&HFFFFFFFE", "D")
 MsgBox "Return from RegWriteValue for large DWord: " & r

r = Sys.RegGetValue("HKCU", "Software\TestM", "DVal1", a1, vType)
  s = "Return from DWrod RegGetValue: " & r & vbcrlf
  s = s & "Value: " & a1 & vbcrlf
  s = s & "Type: " & vtype
  msgbox s

r = Sys.RegGetValue("HKCU", "Software\TestM", "DVal2", a1, vType) 
  s = "Return from DWrod RegGetValue: " & r & vbcrlf
 '--   if (a1 < 0) then a1 = Hex(a1)  '-- method to return hex DWord string rather than decimal number.
  s = s & "Value: " & a1 & vbcrlf
  s = s & "Type: " & vtype
  msgbox s


Set Sys = Nothing