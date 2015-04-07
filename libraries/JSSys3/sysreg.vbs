Dim Sys, sPath, sVal, r, Resp

Resp = MsgBox("Click YES to create Registry keys and values. Click NO to delete them.", 35)
   If r = 2 Then
      WScript.Quit
   End If
   
Set Sys = CreateObject("JSSys3.Ops")
  sPath = "Software\TestSys"
  
If Resp = 6 Then
  
         r = Sys.RegWriteValue("HKCU", sPath, "Value1", "apple", "S")
           If r <> 0 Then MsgBox "write value apple return: " & r
           
         r = Sys.RegWriteValue("HKCU", sPath & "\Key2\Key3", "value3", "orange", "S")
          If r <> 0 Then 
              MsgBox "write value orange return: " & r
          Else
              MsgBox "ok"
              Set Sys = Nothing
          End If
 
Else '--delete key:
 
    r = Sys.RegDeleteVal("HKCU", sPath, "Value1")
      If r <> 0 Then MsgBox "delete value return: " & r
      
  Resp = MsgBox("Value deleted. Delete all new keys?", 36)

    If Resp = 6 Then
        r = Sys.RegDeleteSubkey("HKCU", "Software", "TestSys")
          If r <> 0 Then MsgBox "delete key return: " & r
    End If
    
  Set Sys = Nothing
 End If      
