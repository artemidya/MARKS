Public Sub zProtected_RaiseEvent_OnRecive(ByRef bytes() As Byte, ByRef BlocksSocket As BlocksSocket)
    If zProtected_MeHaweTheParentFlag Then
        Parent.zProtected_RaiseEvent_OnRecive bytes, BlocksSocket
    Else
        RaiseEvent OnRecive(bytes, BlocksSocket)
        LastRefresTiks = GetTickCount
    End If
End Sub
