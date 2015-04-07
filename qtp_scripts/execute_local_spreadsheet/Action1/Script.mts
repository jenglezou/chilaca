' Local QTP execution of TA Framework

Dim sLocalSpreadSheetPath 'path to local spreadsheet
Dim sQtpDataPath

'for edit
sLocalSpreadSheetPath = "c:\VBSFramework\tests\HelloWorldTest.xls"

'path to Data spreadsheet or env_configuration shoud be same as in ALM 
'default is c:\VBSFramework\Temp

sQtpDataPath = "c:\VBSFramework\tests\" 'its necessary to put backslash at the end
ExecuteFile "c:\VBSFramework\VBSFrameworkLauncher.vbs"


