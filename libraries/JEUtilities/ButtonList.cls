VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ButtonList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Private arrButtonList() As String

Public Caption
Public Description
Public HorizontalPosition
Public VerticalPosition
Public Orientation

Public Sub Reset()
    Dim i
    Dim iNumberButtons As Integer

    ReDim arrButtonList(0)
    Caption = "Select Button"
    Description = "Options"
    HorizontalPosition = "Centre"
    VerticalPosition = "Centre"
    
    'frmButtonList.Caption = Caption
    'frmButtonList.txtDescription = Description
    
    iNumberButtons = frmButtonList.Button.UBound
    
    For i = 1 To iNumberButtons Step 1
        Unload frmButtonList.Button(i)
    Next i
End Sub

Public Sub Show()
    Dim i
    Dim iNumberButtons As Integer

    frmButtonList.Caption = Caption
    frmButtonList.txtDescription = Description

    iNumberButtons = UBound(arrButtonList)
    
    frmButtonList.Button(0).Caption = arrButtonList(0)
    frmButtonList.Button(0).Enabled = True
    frmButtonList.Button(0).Visible = True
    
    Orientation_High (iNumberButtons)
    'Orientation_Wide (iNumberButtons)
    
    frmButtonList.Hide
    frmButtonList.Show vbModal
End Sub

Public Sub AddButton(Optional ByVal sButton As String)
    Dim i
    
    i = UBound(arrButtonList) 'Current upper bound
    ReDim Preserve arrButtonList(i + 1)
    arrButtonList(i) = CStr(sButton)
End Sub

Public Function SelectedButton() As String
    SelectedButton = frmButtonList.sSelectedButton
End Function

Private Sub Orientation_High(iNumberButtons As Integer)
    Dim i
    
    For i = 1 To iNumberButtons - 1 Step 1
        Load frmButtonList.Button(i)
        frmButtonList.Button(i).Caption = arrButtonList(i)
        frmButtonList.Button(i).Enabled = True
        frmButtonList.Button(i).Visible = True
        'frmButtonList.Button(i).Top = 180 + 615 * i
        frmButtonList.Button(i).Top = 420 + 615 * i
    Next i

    frmButtonList.Height = iNumberButtons * 615 + 1500
    frmButtonList.CancelButton.Top = frmButtonList.Height - 1000
    
    SetFormHorizontalPosition frmButtonList, HorizontalPosition
    SetFormVerticalPosition frmButtonList, VerticalPosition
    frmButtonList.Move frmButtonList.Left, frmButtonList.Top
End Sub

Private Sub Orientation_Wide(iNumberButtons As Integer)
    Dim i
    
    
    frmButtonList.Button(0).Top = 480
    frmButtonList.Button(0).Width = 1215
    
    For i = 1 To iNumberButtons - 1 Step 1
        On Error Resume Next
        Unload frmButtonList.Button(i)
        On Error GoTo 0
        Load frmButtonList.Button(i)
        frmButtonList.Button(i).Caption = arrButtonList(i)
        frmButtonList.Button(i).Enabled = True
        frmButtonList.Button(i).Visible = True
        frmButtonList.Button(i).Top = 480
        frmButtonList.Button(i).Left = 100 + (1315) * i
        frmButtonList.Button(i).Width = 1215
    Next i
    
    frmButtonList.Height = 1440
    frmButtonList.Width = (iNumberButtons + 1) * 1315 + 200
   
    frmButtonList.CancelButton.Width = 1215
    frmButtonList.CancelButton.Left = iNumberButtons * 1315 + 100
    frmButtonList.CancelButton.Top = 480
    
    SetFormHorizontalPosition frmButtonList, HorizontalPosition
    SetFormVerticalPosition frmButtonList, VerticalPosition
    frmButtonList.Move frmButtonList.Left, frmButtonList.Top
End Sub

