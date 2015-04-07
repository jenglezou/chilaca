'--RegComp2.vbs - this script can be used to register or unregister a component.
'--the method used here allows for drag-and-drop if you have at least version 5.1
'--of the Windows Script Host. You can just put the component
'--in the system folder and then drag it onto this script's file icon. You can also
'--double-click this file to use an inputbox method.
'--The method used to register is silent. RegSvr32 does not show a confirmation message,
'--making this a good method for remote install when no one is there to click OK.

'--To remote install for someone who doesn't know how to work it:

'--1) Make a 2-line BAT file, reg.bat, that calls wscript.exe:

'--        @echo off
'--         wscript.exe c:\windows\temp\reg.vbs

'--2)  Make a script named reg.vbs that will move your script from TEMP to the system folder.
'-- 3) Use code like below to call RegSvr32 to register the component.
'-- 4) Put the component, the BAT file and reg.vbs in a self-executing zip that unloads to TEMP
'--and executes reg.bat.

'-------------------------------------------------------------------------------

Dim sh, arg, r, sys, fso
Set sh = CreateObject("WScript.Shell")
Set fso = CreateObject("Scripting.FileSystemObject")
   

     If wscript.arguments.count = 0 then
           arg = inputbox("This script can register system files in Windows. Enter path of file to register or unregister.", "Register or Unregister", "C:\Windows\System")
     else
           arg = wscript.arguments.item(0)
     end if

  If fso.FileExists(arg) = false then
     msgbox "The path is wrong. No such file.", 64, "Wrong Path"
    
  end if
   sys = fso.GetSpecialFolder(1)

   r = msgbox("Click YES to register." & vbcrlf & "Click NO to Unregister.", 35, "Register or Unregister?")
     if r = 2 then
        wscript.Quit
     elseif r = 6 then
   '--------  this is the only line you need to register: Regsvr path [space] /S [space] DLL path   ------------
      '-- silent method:  sh.Run(sys & "\regsvr32.exe /S " & arg)
         sh.Run(sys & "\regsvr32.exe " & arg)
     else
       '--silent method:   sh.Run(sys & "\regsvr32.exe /S /U " & arg)
       sh.Run(sys & "\regsvr32.exe /U " & arg)
     end if
    'msgbox "Done."
