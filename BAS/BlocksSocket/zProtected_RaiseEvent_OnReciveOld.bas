Public Sub zProtected_RaiseEvent_OnReciveOld(ByRef bytes() As Byte, ByRef BlocksSocket As BlocksSocket)
    If zProtected_MeHaweTheParentFlag Then
        Parent.zProtected_RaiseEvent_OnReciveOld bytes, BlocksSocket
    Else
        RaiseEvent OnReciveOld(bytes, BlocksSocket)
        LastRefresTiks = GetTickCount
    End If
End Sub
