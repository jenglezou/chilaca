Dim sLocalSpreadSheetPath
Dim sQtpResultPath '  path where result spreadsheet will be stored
Dim sQtpDataPath '  path wher are data spreadsheets stored

Dim bRunner:bRunner=True

sLocalSpreadSheetPath = TestArgs("TestFile")
sQtpResultPath = TestArgs("ResultFolder")
sQtpDataPath =TestArgs("DataFolder")

'sLocalSpreadSheetPath = "c:\VBSFramework\tests\test\HelloWorldTest.xls"

ExecuteFile "c:\VBSFramework\VBSFrameworkLauncher.vbs"


