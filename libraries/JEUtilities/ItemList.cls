VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ItemList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Private arrItemList() As String

Public Caption
Public Description
Public HorizontalPosition
Public VerticalPosition

Public Sub Reset()
    ReDim arrItemList(0)
    Caption = "Select Item"
    Description = "Options"
    HorizontalPosition = "Centre"
    VerticalPosition = "Centre"

    'frmItemList.Caption = Caption
    'frmItemList.txtDescription = Description
    
    frmItemList.List1.Clear
End Sub

Public Sub Show()
    Dim i
    
    frmItemList.Caption = Caption
    frmItemList.txtDescription = Description
    For i = 0 To UBound(arrItemList) - 1 Step 1
        frmItemList.List1.Additem arrItemList(i)
    Next i
    
    SetFormHorizontalPosition frmItemList, HorizontalPosition
    SetFormVerticalPosition frmItemList, VerticalPosition
    frmItemList.Move frmItemList.Left, frmItemList.Top
    
    frmItemList.Hide
    frmItemList.Show vbModal
End Sub

Public Sub Additem(Optional ByVal sItem As String)
    Dim i
    
    i = UBound(arrItemList)
    ReDim Preserve arrItemList(i + 1)
    arrItemList(i) = CStr(sItem)
End Sub

Public Function SelectedItem() As String
    SelectedItem = frmItemList.sSelectedItem
End Function

Public Function ButtonPressed() As Integer
    ButtonPressed = frmItemList.vbButtonPressed
End Function


