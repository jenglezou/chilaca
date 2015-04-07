Private Sub Form_Load()

    test '"hello"
    
    End
End Sub

Private Sub test(Optional ByVal sMessage As Variant)

    MsgBox CStr(sMessage)
    
    If IsNull(sMessage) Then
        MsgBox "isnull:" & CStr(sMessage)
    End If
    
    If IsEmpty(sMessage) Then
        MsgBox "isempty:" & CStr(sMessage)
    End If
    
    If Len(CStr(sMessage)) = 0 Then
        MsgBox "len = 0:" & CStr(sMessage)
    End If

End Sub
