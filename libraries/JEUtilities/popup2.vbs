on error resume next
Set oPopUp = CreateObject("JEUtilities.PopUp")

if err.number <> 0 then msgbox "error here"
'Check again because it could still fail to be created because the DLL/EXE server may not be loaded
If Not oPopUp Is Nothing Then
	msgbox "trying to show popup"
	oPopUp.WindowTitle = "Progress information"
	oPopUp.ShowProgressBar = false
	oPopUp.Timeout = 10
	oPopUp.Message = "Message"
	'iSeconds = 101
	oPopUp.ShowMsg "hello"
else
	msgbox "Error"
End If
On Error GoTo 0
