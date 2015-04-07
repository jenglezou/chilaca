'-- demo of browsing dialogues.
 
Dim Sys, r

Set Sys = CreateObject("JSSys3.Ops")

MsgBox "File browsing demo. There are 5 examples: Open. Open Multi. Save. Color. BrowseForFolder."

 r = Sys.OpenDlg("Open File", "bmp", "C:\")
 
MsgBox r
WScript.sleep 500

r = Sys.OpenMultiDlg("Select any number of files", "", "")

MsgBox r
WScript.sleep 500

r = Sys.SaveDlg("Save file", "txt")

MsgBox r
WScript.sleep 500

r = Sys.ColorDlg

MsgBox r
WScript.sleep 500

r = Sys.BrowseForFolder("Select a folder.")
msgbox r

Set Sys = Nothing
                    
     