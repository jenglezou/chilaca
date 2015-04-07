on error resume next
set oPopUp = createobject("JEUtilities.PopUp")
if oPopUp is nothing then 
	msgbox "no object"
else
	oPopUp.Timeout = 10
	oPopUp.ShowMsg "Processing please wait ....."
	msgbox "Click OK to close the PopUp"
	'oPopUp.CloseMsg	
	oPopUp.ShowMsg "Still Processing please wait ....."
	msgbox "Click OK to close the PopUp"
	'oPopUp.CloseMsg	
	oPopUp.ShowProgressBar = True
	oPopUp.ShowMsg "Progressbar forever. Click ok to close"
end if


