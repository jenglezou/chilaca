Set WshShell = CreateObject("WScript.Shell")
Set WshSystemEnv = WshShell.Environment("SYSTEM")
WshSystemEnv("LSHOST") = "myLicenseHost"