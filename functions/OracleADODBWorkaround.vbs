Option Explicit 
Dim Args, sConn, sQuery, iMaxOutput, Arg, Name, Value
'msgbox "In Here"
set Args = Wscript.Arguments
'msgbox Args.count
For each Arg in Args
	Name = split(Arg, ":")(0)
	Value = split(Arg, ":")(1)
	select case uCase(Name)
	case "CONN" 
		sConn = Value
	case "QUERY"
		sQuery = Value
	case "MAXOUTPUT"
		iMaxOutput = Value
	end select
Next

iMaxOutput = 1024	

Dim oConn, oRS
Set oConn = CreateObject("ADODB.Connection")
Set oRS = CreateObject("ADODB.Recordset")

oConn.Open sConn
oRS.Open sQuery, oConn

Dim sOutput, i
Do while not oRS.EOF
	for i = 0 To oRS.Fields.Count - 1
		sOutput = sOutput & "," & oRS.Fields(i).name & "=" & oRS.Fields(i).Value	
	Next
	if len(sOutput) > iMaxOutput then exit Do
	sOutput = sOutput & vbNewLine
	oRS.MoveNext
Loop

dim wshShell, oIn, oExec
Set wshShell=createobject("wscript.shell")

Set oExec = wshShell.exec("clip")
Set oIn = oExec.StdIn

oIn.Write(sOutput)

oIn.Close

While oExec.Status= 0
	wscript.sleep 100
Wend 

Set oIn = nothing
Set oExec = Nothing
Set wshShell = Nothing
	
Set oRS = Nothing
Set oConn = Nothing
