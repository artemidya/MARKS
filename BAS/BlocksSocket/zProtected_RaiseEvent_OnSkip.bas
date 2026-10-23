Public Sub zProtected_RaiseEvent_OnSkip(ByRef BlocksSocket As BlocksSocket)
    If zProtected_MeHaweTheParentFlag Then
        Parent.zProtected_RaiseEvent_OnSkip BlocksSocket
    Else
        RaiseEvent OnSkip(BlocksSocket)
        LastRefresTiks = GetTickCount
    End If
End Sub
