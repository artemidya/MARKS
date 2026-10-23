Public Sub zProtected_RaiseEvent_OnStatisticChange(ByRef BlocksSocket As BlocksSocket)
    If zProtected_MeHaweTheParentFlag Then
        Parent.zProtected_RaiseEvent_OnStatisticChange BlocksSocket
    Else
        RaiseEvent OnStatisticChange(BlocksSocket)
        LastRefresTiks = GetTickCount
    End If
End Sub
