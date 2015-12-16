REM *****************   This can be run in a command window started as ADMINISTRATOR *****************************
c:\windows\microsoft.net\framework\v2.0.50727\regasm.exe ..\libraries\Selenium.dll /tlb:..\libraries\selenium.tlb /codebase
start /MIN c:\windows\syswow64\cscript .\LocalCMDLauncher.vbs ..\tests\SeleniumChromeTest.xls
