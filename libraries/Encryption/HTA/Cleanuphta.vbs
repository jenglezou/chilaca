strComputer = "."
Set objWMIService = GetObject("winmgmts:" & "{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")
Set colProcessList = objWMIService.ExecQuery ("Select * from Win32_Process Where Name = 'mshta.exe'")
For Each objProcess in colProcessList
	objProcess.Terminate()
Next
