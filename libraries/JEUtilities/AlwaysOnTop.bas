Attribute VB_Name = "AlwaysOnTop"
Option Explicit

Private Declare Function CreateRectRgn Lib "gdi32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Declare Function CombineRgn Lib "gdi32" (ByVal hDestRgn As Long, ByVal hSrcRgn1 As Long, ByVal hSrcRgn2 As Long, ByVal nCombineMode As Long) As Long
Private Declare Function SetWindowRgn Lib "user32" (ByVal hWnd As Long, ByVal hRgn As Long, ByVal bRedraw As Long) As Long

Private Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, _
    ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, _
    ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long

Public Declare Function FlashWindow Lib "user32" (ByVal hWnd As Long, ByVal bInvert As Long) As Long

Const SWP_NOSIZE = &H1
Const SWP_NOMOVE = &H2
Const SWP_SHOWWINDOW = &H40
Const HWND_NOTOPMOST = -2
Const HWND_TOPMOST = -1

' Set a form always on the top.
'
' the form can be specified as a Form or object
' or through its hWnd property
' If OnTop=False the always on the top mode is de-activated.
Public Sub SetAlwaysOnTopMode(hWndOrForm As Variant, Optional ByVal OnTop As Boolean = _
    True)
    Dim hWnd As Long
    ' get the hWnd of the form to be move on top
    'MsgBox "SetAlwaysOnTopMode: " & OnTop
    If VarType(hWndOrForm) = vbLong Then
        hWnd = hWndOrForm
    Else
        hWnd = hWndOrForm.hWnd
    End If
    SetWindowPos hWnd, IIf(OnTop, HWND_TOPMOST, HWND_NOTOPMOST), 0, 0, 0, 0, _
        SWP_NOMOVE Or SWP_NOSIZE Or SWP_SHOWWINDOW
End Sub

Public Sub GlassifyForm(frm As Form)
Const RGN_DIFF = 4
Const RGN_OR = 2

Dim outer_rgn As Long
Dim inner_rgn As Long
Dim wid As Single
Dim hgt As Single
Dim border_width As Single
Dim title_height As Single
Dim ctl_left As Single
Dim ctl_top As Single
Dim ctl_right As Single
Dim ctl_bottom As Single
Dim control_rgn As Long
Dim combined_rgn As Long
Dim ctl As Control

    If frm.WindowState = vbMinimized Then Exit Sub

    ' Create the main form region.
    wid = frm.ScaleX(frm.Width, vbTwips, vbPixels)
    hgt = frm.ScaleY(frm.Height, vbTwips, vbPixels)
    outer_rgn = CreateRectRgn(0, 0, wid, hgt)

    border_width = (wid - frm.ScaleWidth) / 2
    title_height = hgt - border_width - frm.ScaleHeight
    inner_rgn = CreateRectRgn( _
        border_width, _
        title_height, _
        wid - border_width, _
        hgt - border_width)

    ' Subtract the inner region from the outer.
    combined_rgn = CreateRectRgn(0, 0, 0, 0)
    CombineRgn combined_rgn, outer_rgn, _
        inner_rgn, RGN_DIFF

    ' Create the control regions.
    For Each ctl In frm.Controls
        On Error Resume Next
        If ctl.Container Is frm Then
'        If ctl.Visible = True Then
            ctl_left = frm.ScaleX(ctl.Left, frm.ScaleMode, vbPixels) _
                + border_width
            ctl_top = frm.ScaleX(ctl.Top, frm.ScaleMode, vbPixels) _
                + title_height
            ctl_right = frm.ScaleX(ctl.Width, frm.ScaleMode, vbPixels) _
                + ctl_left
            ctl_bottom = frm.ScaleX(ctl.Height, frm.ScaleMode, vbPixels) _
                + ctl_top
            control_rgn = CreateRectRgn( _
                ctl_left, ctl_top, _
                ctl_right, ctl_bottom)
            CombineRgn combined_rgn, combined_rgn, _
                control_rgn, RGN_OR
        End If
        On Error GoTo 0
    Next ctl

    ' Restrict the window to the region.
    SetWindowRgn frm.hWnd, combined_rgn, True
End Sub

Public Sub SetFormHorizontalPosition(frmForm As Form, sHorizontalPosition As Variant)
    'MsgBox sHorizontalPosition
    Select Case UCase(Mid(sHorizontalPosition, 1, 1))
    Case "C"
        frmForm.Left = (Screen.Width - frmForm.Width) / 2
    Case "L"
        frmForm.Left = 1
    Case "R"
        frmForm.Left = Screen.Width - frmForm.Width
    Case Default
        frmForm.Left = sHorizontalPosition
    End Select
End Sub

Public Sub SetFormVerticalPosition(frmForm As Form, sVerticalPosition As Variant)
    'MsgBox sVerticalPosition
    Select Case UCase(Mid(sVerticalPosition, 1, 1))
    Case "C"
        frmForm.Top = (Screen.Height - frmForm.Height) / 2
    Case "T"
        frmForm.Top = 1
    Case "B"
        frmForm.Top = Screen.Height - frmForm.Height
    Case Default
        frmForm.Top = sVerticalPosition
    End Select
End Sub

'Private Sub Form_Load()
'    Me.ScaleMode = 3
'End Sub
'
'Private Sub Form_Resize()
'    GlassifyForm Me
'End Sub

