'-- Updated 12-12-02.
'-- >>> Requires JSSys3.DLL - Not JSSys.DLL.
'-- CPU load script.
'-- This script uses the JSSys3.dll comonent to 
'-- Get the current CPU usage For Win9x/ME.
'-- It first signals Windows to start monitoring the CPU usage
'--by reading the value in "PerfStats\StartStat".
'--Then it reads the usage every 3 seconds and reports
'--it Until CANCEL is clicked.
'--You can see changes in usage by opening and closing programs
'--While the script is running.
'-- When CANCEL is clicked the script reads the "PerfStats\StopStat"
'-- CPUUsage value to stop monitoring.

'-- This will Not work on NT/2000/XP.
'-- RegGetValue does Not read from HKEY_PERFORMANCE_DATA.

'--  JSSys3.dll is Not required For this Function but you Do need
'--  a means to overcome the inability of VBS to read from
'--  HKEY_DYN_DATA



'-- Win95 note: This script should work with Win95 but in tests
'-- the RegGetValue Call failed. I don't know the reason For that.
'--------------------------------------------------------
Dim SH, Sys, s, StartD, GetD, StopD, sType, num, r, i

Set Sys = createobject("JSSys3.Ops")

StartD = "PerfStats\StartStat"
GetD = "PerfStats\StatData"
StopD = "PerfStats\StopStat"

     '--start statistic reporting by reading the startstat value:

     i = Sys.RegGetValue("HKDD", StartD, "KERNEL\CPUUsage", s, sType)     
      
  If  i = 0 Then      '--means read was successful.
       MsgBox "CPU usage will be reported every 3 seconds. Click Cancel when you want it to stop.", 64, "CPU usage"
  Else
       MsgBox "Unable to start performance monitoring.", 64, "CPU usage"
       WScript.Quit
  End If
            
Do
    i = Sys.RegGetValue("HKDD", GetD, "KERNEL\CPUUsage", s, sType)    
          If (i = 0) and (sType = "B") Then
               num = s(0)
          End If
    r = MsgBox("Current CPU usage is " & num & " percent.", 65, "CPU usage")
         If r = 2 Then
             Exit Do
         End If
      
    WScript.Sleep 3000
Loop

                   '--stop data reporting. This is needed to prevent Windows from tracking
                   '--stats Until shutdown:

                i = Sys.RegGetValue("HKDD", StopD, "KERNEL\CPUUsage", s, sType)     
         
          
             
  