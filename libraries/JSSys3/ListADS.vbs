'-- List ADS stream files. Drop a folder (or a drive icon) onto script.
'  A list of ADS files will be saved on C drive.


Dim SOps, sFils, Arg, FSO, sDenied, TS
Set FSO = CreateObject("Scripting.FileSystemObject")

Arg = WScript.Arguments(0)
  If FSO.FolderExists(Arg) = False Then
     MsgBox "A folder must be dropped onto script.", 64
     Set FSO = Nothing
     WScript.Quit
   End If

Set SOps = CreateObject("JSSys3.StreamOps")

FindStreams Arg

Set TS = FSO.CreateTextFile("C:\ADSList.txt", True)
  TS.write "Streams in " & Arg & ":" & vbCrLf & vbCrLf & sFils & vbCrLf & vbCrLf & "Folders where access was denied:" & vbCrLf & vbCrLf & sDenied
   TS.Close
 Set TS = Nothing

MsgBox "ADS list saved as C:\ADSList.txt"

DropIt

'------------------------------------------
Sub FindStreams(sFolPath)
  Dim LNum, oFol, oFols, oFol1, oFils, oFil, AFils, ASz, i2
    On Error Resume Next  ' there may be "permission denied" problems.
    '-- get streams for folder itself.
     LNum = SOps.ListStreams(sFolPath, AFils, ASz)
   If LNum > 0 Then
        For i2 = 0 to LNum - 1
          sFils = sFils & sFolPath & ":" & AFils(i2) & " * " & ASz(i2) & vbCrLf
        Next
    End If

  Set oFol = FSO.GetFolder(sFolPath)
    '-- get streams for all files in folder.
  Set oFils = oFol.Files
    If oFils.count > 0 Then
         If Err.number <> 0 Then sDenied = sDenied & sFolPath & vbCrLf: Err.clear
        For Each oFil in oFils
           LNum = SOps.ListStreams(oFil.Path, AFils, ASz)
           If LNum > 0 Then
              For i2 = 0 to LNum - 1
                sFils = sFils & oFil.path & ":" & AFils(i2) & "*" & ASz(i2) & vbCrLf
              Next
           End If
        Next   
     End If      
  Set oFils = Nothing
    '-- recursively get subfolders/files.
  Set oFols = oFol.SubFolders
    If oFols.count > 0 Then
      For Each oFol1 in oFols
        FindStreams oFol1.Path
      Next
    End If  
  Set oFols = Nothing  
 Set oFol = Nothing
End Sub

Sub DropIt()
  Set SOps = Nothing
  Set FSO = Nothing
  WScript.quit
End Sub
 