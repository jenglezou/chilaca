REM STARTUP - Add the registry entry to always run the startup script to configure QTP Remote Host User Settings
Regedit /s c:\VBSFramework\setup\Startup\AddToStartupQTPRemoteHost.reg

REM Run the Startup Configuration now 
c:\VBSFramework\setup\Startup\RunOnStartupQTPRemoteHost.bat