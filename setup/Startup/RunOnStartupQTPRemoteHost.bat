REM In case a windows update screws up the qtp com settings ...
REM "%ProgramFiles%\HP\QuickTest Professional\bin\QTAutomationAgent.exe" /regserver

REM Add LSHOST into the system environment variables if it isn't set
IF "%LSHOST%"=="" c:\VBSFramework\setup\Startup\SetLSHOST.vbs

REM Configure Power Settings
powercfg /import desktop /file c:\VBSFramework\setup\Startup\desktop.pwr
powercfg /setactive desktop

REM Region Setting - set to US (0409)
Rundll32 shell32,Control_RunDLL intl.cpl,,/f:"c:\VBSFramework\setup\Startup\RegionUK-0809.ini"

REM VBScript to set registry entries for QTP, Screensaver and Wallpaper
c:\VBSFramework\setup\Startup\ConfigureRegistrySettingsQTPRemoteHost.vbs

Regedit /s c:\VBSFramework\setup\Startup\TortoiseSVNSettings.reg