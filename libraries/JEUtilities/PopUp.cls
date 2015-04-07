VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "PopUp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Public WindowTitle As String
Public ShowProgressBar As Boolean
Public TimeOut As Long
Public Message As String

Public Sub ShowMsg(Optional ByVal sMessage As String)

    If TimeOut > 65 Then TimeOut = 65
    frmPopUp.bShow = True
    frmPopUp.Caption = "Automation Information"
    If Len(WindowTitle) > 0 Then frmPopUp.Caption = WindowTitle
    
    frmPopUp.Label1 = CStr(sMessage)
    If Len(CStr(sMessage)) = 0 Then frmPopUp.Label1 = Message
    
    frmPopUp.Timer1.Interval = TimeOut * 1000
    frmPopUp.Label2 = frmPopUp.Timer1.Interval
    frmPopUp.ProgressBar1.Value = 0
    frmPopUp.ProgressBar1.Visible = False
    If ShowProgressBar = True Then
        frmPopUp.ProgressBar1.Visible = True
    End If
    
    frmPopUp.Hide
    frmPopUp.Show vbModeless

End Sub

Public Sub CloseMsg()
    If frmPopUp.bShow = True Then
        frmPopUp.Timer1.Interval = 0
        frmPopUp.bShow = False
        Unload frmPopUp
    End If
End Sub
