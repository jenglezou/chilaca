'-- demo script For GetRegFromCLSID and GetRegFromProgID
 '-- Function returns an array of ubound 5 whether it succeeds or fails.
 '--error code is in Array(0) . 
 '--Return: 0 - item found and array holds data'
  '--          1 - item Not found.
  
Dim Sys, a, s, sRes
Set Sys = CreateObject("JSSys3.Ops")

s = "{F738999A-E160-11D6-B5C6-C78C22D19941}"
a = Sys.GetRegFromCLSID(s)
 If a(0) = 1 Then
   MsgBox "No listing for " & s
 Else
    sRes = "CLSID: " & s & vbcrlf
    sRes = sRes & "ProgID: " & a(2) & vbcrlf
    sRes = sRes & "VersionIndependentProgID: " & a(3) & vbcrlf
    sRes = sRes & "InprocServer32: " & a(4) & vbcrlf
    sRes = sRes & "LocalServer32: " & a(5) & vbcrlf
    MsgBox sRes
 End If
 
 s = "JSSys3.Ops"
 a = Sys.GetRegFromProgID(s)   
   If a(0) = 1 Then
      MsgBox "No listing for " & s
   Else
      sRes = "CLSID: " & a(1) & vbcrlf
      sRes = sRes & "ProgID: " & s & vbcrlf
      sRes = sRes & "VersionIndependentProgID: " & a(3) & vbcrlf
      sRes = sRes & "InprocServer32: " & a(4) & vbcrlf
      sRes = sRes & "LocalServer32: " & a(5) & vbcrlf
      MsgBox sRes
   End If