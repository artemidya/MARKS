Public Sub zProtected_RaiseEvent_OnConnect(ByRef BlocksSocket As BlocksSocket)
    If zProtected_MeHaweTheParentFlag Then
        Parent.zProtected_RaiseEvent_OnConnect BlocksSocket
    Else
        RaiseEvent OnConnect(BlocksSocket)
        LastRefresTiks = GetTickCount
    End If
End Sub
