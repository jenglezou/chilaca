'--___________________________________________________
'--INI FILE OPERATIONS WITH JSSys3.IniEdit
'--___________________________________________________
'--
'--To use in VBScript:
'--
'--    Dim ObjVar 
'--    Set ObjVar = CreateObject("JSSys3.IniEdit")
'--__________________________________________________


'---------------BEGIN SCRIPT--------------------------------------------------------------
'-----NOTE: this script requires that the enclosed file, blah.ini, be on the Desktop-------
'-----alternatively, you can change the script paths.-------------------------------------------
'-------------------------------------------------------------------------------------------
option explicit
 Dim ob, v, s1, a, i, bool, s, SH
 Set ob = CreateObject("JSSys3.IniEdit")
Set SH = CreateObject("WScript.Shell")
 s = SH.SpecialFolders("Desktop")
s = s & "\blah.ini"

'-----------------get all key=value pairs in section.---------------------------------
    '--gets all keys, with their values, in the [blah] section.

  v = ob.GetIniSectionVals(s, "blah")
    If v <> "" Then
       If InStr(1, v, Chr(0), 0) = 0 Then  '--if no Chr(0) then there's only one key-value pair.
          s1 = v
       Else
         a = Split(v, Chr(0))
           For i = 0 To UBound(a)
              s1 = s1 & a(i) & vbCrLf
           Next
       End If
      MsgBox "Key-Value pairs in the BLAH section:" & vbcrlf & s1
    Else
      MsgBox "failed"
    End If
s1 = ""

'-------------------get all keynames in section.------------------------------
    '--this will get only the key names, without the values.

 v = ob.GetIniSectionkeys(s, "glooie")
    If v <> "" Then
        If InStr(1, v, Chr(0), 0) = 0 Then     '--if no Chr(0) then there's only one key.
           s1 = v
        Else
         a = Split(v, Chr(0))
           For i = 0 To UBound(a)
              s1 = s1 & a(i) & vbCrLf
           Next
       End If
      MsgBox "Keynames in the GLOOIE section:" & vbcrlf &  s1
    Else
      MsgBox "failed"
    End If
s1 = ""

'-----------------------get all section names in file.------------------------------------------------
 
 v = ob.GetIniSectionNames(s)
    If v <> "" Then
      If InStr(1, v, Chr(0), 0) = 0 Then   '--if no Chr(0) then there's only one section.
        s1 = v
      Else
        a = Split(v, Chr(0))
          For i = 0 To UBound(a)
            s1 = s1 & a(i) & vbCrLf
          Next
      End If
       MsgBox "All section names in the BLAH.INI file:" & vbcrlf & s1
    Else
       MsgBox "failed"
    End If
s1 = ""

'----------------get one value.------------------------------------------------
'----------------get the value for the "sticky" key in the [glooie] section.

v = ob.GetIniVal(s, "glooie", "sticky")
   If v <> "" Then
        MsgBox "The value for STICKY in the GLOOIE section is: " & v
   else
        msgbox "failed"
   End If

'-----------------write a section.-----------------------------------------------------
    '--this will write a section called [NewSection] with 3 values.

  s1 = "new1=value1" & Chr(0) & "new2=value2" & Chr(0) & "new3=value3" & Chr(0)
 
   bool = ob.WriteIniSection(s, "NewSection", s1)
      If bool = True Then
          MsgBox "A new section has just been written called NewSection and 3 values have been added."
      else
          msgbox "failed"
      end if

 '-------------------------write a value.-------------------------------------------------
          '--this adds a 4th value to [NewSection] : key4=value4

 bool = ob.WriteIniVal(s, "NewSection", "key4", "value4")
   If bool = True Then
      MsgBox "A 4th value has been added to the NewSection section, called KEY4, with a value of VALUE4."
  else
      msgbox "failed"
  End If

'--to delete the "key4" key:
 '-- bool = ob.WriteIniVal(s, "NewSection", "key4", "null")

'--to delete the [NewSection] section:
'--  bool = ob.WriteIniVal(s, "NewSection", "null", "null")
 