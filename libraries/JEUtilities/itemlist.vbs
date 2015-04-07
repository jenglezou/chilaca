on error resume next
set oItemList = createobject("JEUtilities.itemlist")
if oItemList is nothing then 
	msgbox "no object"
else
	oItemList.Reset
	oItemList.Caption = "Select string"
	oItemList.AddItem 
	oItemList.AddItem x
	oItemList.AddItem 2.1
	oItemList.AddItem "string 1"
	oItemList.AddItem "string 2"
	oItemList.AddItem "string 4"
	oItemList.AddItem "string 3"
	oItemList.AddItem "string XXX"
	oItemlist.show
	msgbox oItemList.SelectedItem
	oItemList.Caption = "Select string2"
	oItemlist.show
	msgbox oItemList.SelectedItem
	oItemList.Reset
	oItemList.AddItem "string 1"
	oItemList.AddItem "string XXX"
	oItemlist.show
	msgbox oItemList.SelectedItem
	
end if


