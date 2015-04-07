'--Hardware info. script. - For Win 9x/ME only.
 '--updated 12-12-02 For JSSys3.DLL. Will Not work with JSSys.DLL.
'-- This script uses the JSSys3.dll component to find information
'-- in the Registry about installed hardware. 
'-- The path to this info is very circuitous but comments have
'-- been added to the script to document it.
'-- The script will Get the information and write it to
'-- a text file: C:\HardwareInfo.txt.

'-- JSSys3.dll is Not required For this Function but you Do need
'--a means to overcome 2 limitations of VBS:
'--  1) the inability to enumerate subkeys.
'--  2) the inability to read from HKEY_DYN_DATA

'--------------------------------------------------------
Dim FSO, SH, Sys, AList, i, ts, pDriv, i1, sKey, sPath, sType, i2

Set FSO = CreateObject("Scripting.FileSystemObject")
Set SH = CreateObject("WScript.Shell")
Set Sys = createobject("JSSys3.ops")

pDriv = "HKLM\System\CurrentControlSet\Services\Class\"

'--Get array of key values representing currently installed hardware:

i2 = Sys.RegListKeys("HKDD", "Config Manager\Enum", AList)
   If i2 <> 0 Then
       MsgBox "Failure obtaining Registry Dynamic Data listing.", 64
       WScript.Quit
   End If

'--start a file to record info.:

Set ts = FSO.CreateTextFile("C:\HardwareInfo.txt", True)

 '--go through the array of subkey names in HKDD\Config Manager\Enum to Get info.
'--one problem: vbs doesn't read HKDD. So this routine uses Ops
'-- to Get the "HardwareKey" data string - the one needed to complete path 
'--to driver info. key.

 For i = 0 to ubound(AList)  
     sKey = AList(i)
   
     i2 = Sys.RegGetValue("HKDD", "Config Manager\Enum\" & sKey, "HardWareKey", sPath, sType)
          
     '--If "HardwareKey" value was found, Get info.:
          If i2 = 0 and sType = "S" Then   '--must be capital S.
                Call ListHardwareItem(sPath)
          End If
 Next
 

     ts.Close
     Set ts = Nothing
     MsgBox "Done. Information is recorded in C:\HardwareInfo.txt", 64


'----------------------------------------------------------------------------

Sub ListHardwareItem(s1)     '--Get registry data about given hardware item.
  Dim sListing, s2
 On Error Resume Next
    '-- Each key under HKDD\Config Manager\Enum represents an installed hardware item.
    '-- The code above gets the string value of "HardwareKey" For Each HKDD\Config Manager\Enum
   '--subkey and sends it to this Sub as the s1 parameter.
   '--In this Sub:
    '-- Use the s1 string to Get info from HKLM\Enum. The value For "Driver" points to the path
    '--For driver info.

  '--first Get the info. under HKLM\Enum\{HardWareKey String}\Class:

  sListing = "Class: " & SH.RegRead("HKLM\Enum\" & s1 & "\Class") & VbCrLf
  sListing = sListing & "Device Description: " & SH.RegRead("HKLM\Enum\" & s1 & "\Devicedesc") & VbCrLf
  sListing = sListing & "Manufacturer: " & SH.RegRead("HKLM\Enum\" & s1 & "\Mfg") & VbCrLf
    ts.Write sListing
    sListing = ""

'--Then check For a "driver" value in that key:

 s2 = SH.RegRead("HKLM\Enum\" & s1 & "\Driver")

  '-- If there was a "Driver" value add it to
  '-- HKLM\System\CurrentControlSet\Services\Class\ to Get the driver info:
 
   If s2 <> "" Then
      sListing = "Driver: " & SH.RegRead(pDriv & s2 & "\Driver") & VbCrLf
      sListing = sListing & "Driver Date: " & SH.RegRead(pDriv & s2 & "\DriverDate") & VbCrLf 
      sListing = sListing & "Driver Description: " & SH.RegRead(pDriv & s2 & "\DriverDesc") & VbCrLf 
      sListing = sListing & "Related INF File: " & SH.RegRead(pDriv & s2 & "\InfPath") & VbCrLf 
      sListing = sListing & "Provider: " & SH.RegRead(pDriv & s2 & "\ProviderName") & VbCrLf 
        ts.Write sListing
        sListing = ""
        ts.Write "_____________________" & vbcrlf
   End If
    
End Sub
