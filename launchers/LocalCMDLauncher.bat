REM ---LOCAL CMD LAUNCHER---
REM {PathTo\VBSFrameworkLauncher.vbs [FullPathToSpreadsheet]}
REM start /MIN cscript /X .\VBSFrameworkLauncher.vbs .\tests\HelloWorldTest.xls
rem start /MIN c:\windows\syswow64\cscript .\VBSFrameworkLauncher.vbs .\tests\Examples.xls
start /MIN c:\windows\syswow64\cscript .\LocalCMDLauncher.vbs ..\tests\Examples.xls
REM start /MIN cscript .\VBSFrameworkLauncher.vbs .\tests\Elsi_iTrent_Demo_Scenario_01.xls