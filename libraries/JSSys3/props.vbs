Dim sys, FilOb, s, arg
    Set sys = CreateObject("jssys3.ops")
    
If WScript.arguments.count <> 0 Then
   arg = WScript.arguments(0)
Else
  MsgBox "Drop a file onto the script to get info.", 64
  WScript.Quit
End If  

On Error Resume Next
Set FilOb = sys.GetPEFile(arg)
 If Err.number <> 0 Then
   MsgBox "Error " & Err.number & vbcrlf & "Description: " & Err.description
   Set sys = Nothing
   WScript.Quit
 End If
  
  
s = "CompanyName: " & FilOb.CompanyName & vbcrlf
s = s & "FileVersion: " & FilOb.FileVersion & vbcrlf
s = s & "DateCreated: " & FilOb.DateCreated & vbcrlf
s = s & "DateLastModified: " & FilOb.DateLastModified & vbcrlf
s = s & "ProductVersion: " & FilOb.ProductVersion & vbcrlf
s = s & "ProductName: " & FilOb.ProductName & vbcrlf
s = s & "FileDescription: " & FilOb.FileDescription & vbcrlf
s = s & "Size: " & FilOb.Size & vbcrlf
s = s & "Name: " & FilOb.Name & vbcrlf

MsgBox s

Set FilOb = Nothing
Set sys = Nothing
