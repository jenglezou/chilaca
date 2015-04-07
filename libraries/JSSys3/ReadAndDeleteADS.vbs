'-- ADS demo for JSSys3.dll.
'-- Drop a file or folder onto script. ADS file attachments will be listed, then deleted.

Dim SOps, LNum, i2, AFils, ASz, sFil, sADS, sList, sText, s2

sFil = WScript.Arguments(0)  
If Len(sFil) = 0 Then
   MsgBox "Drop a file with ADS attachments onto script for demo."
   WScript.Quit
End If

Set SOps = CreateObject("JSSys3.StreamOps")
LNum = SOps.ListStreams(sFil, AFils, ASz)

MsgBox LNum & " ADS attachments found with this file/folder."

If LNum > 0 Then
  sList = "ADS files attached to this file/folder:" & vbCrLf
  For i2 = 0 to LNum - 1
     sADS = AFils(i2)
     sList = sList & sADS & vbCrLf & vbCrLf
       If ASz(i2) > 0 Then
         s2 = SOps.ReadStream(sFil & ":" & sADS, 0)
         s2 = Replace(s2, chr(0), "*")
         sList  = sList & s2
       End If  
         SOps.DeleteStream sFil & ":" & sADS
   Next
      MsgBox "This is a list of ADS names and content:" & vbCrLf & vbCrLf & sList
 End If      
     
 

Set SOps = Nothing

MsgBox "If this is the first run, run the script again to confirm that ADS attachements were deleted."