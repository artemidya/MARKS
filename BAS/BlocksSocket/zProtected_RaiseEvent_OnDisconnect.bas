Public Sub zProtected_RaiseEvent_OnDisconnect(ByRef BlocksSocket As BlocksSocket)
    If zProtected_MeHaweTheParentFlag Then
        Parent.zProtected_RaiseEvent_OnDisconnect BlocksSocket
    Else
        RaiseEvent OnDisconnect(BlocksSocket)
        LastRefresTiks = GetTickCount
    End If
End Sub
