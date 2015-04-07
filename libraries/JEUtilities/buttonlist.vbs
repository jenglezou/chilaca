set oButtonList = createobject("JEUtilities.Buttonlist")
if oButtonList is nothing then 
	msgbox "no object"
else
	oButtonList.Reset
	oButtonList.Caption = "Select string"
	oButtonList.Description = "My Description"
	oButtonList.HorizontalPosition = "Right"
	oButtonList.VerticalPosition = "Top"
	oButtonList.AddButton 
	oButtonList.AddButton x
	oButtonList.AddButton 1
	oButtonList.AddButton "string 1"
	oButtonList.AddButton "string 2"
	oButtonList.AddButton "string 4"
	oButtonList.AddButton "string 3"
	oButtonList.AddButton "string XXX"
	oButtonlist.show
	msgbox oButtonList.SelectedButton

'	oButtonList.Reset
	oButtonList.Caption = "Select string2"
'	oButtonList.AddButton "string 1"
'	oButtonList.AddButton "string 2"
'	oButtonList.AddButton "string 4"
'	oButtonList.AddButton "string 3"
'	oButtonList.AddButton "string XXX"
	oButtonlist.show
	msgbox oButtonList.SelectedButton

	oButtonList.Reset
	oButtonList.AddButton "string 1"
	oButtonList.AddButton "string XXX"
	oButtonlist.show
	msgbox oButtonList.SelectedButton
	
end if


