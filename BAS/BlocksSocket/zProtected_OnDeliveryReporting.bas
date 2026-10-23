Public Sub zProtected_OnDeliveryReporting(BytesDelivered As Long, BytesSize As Long, ByRef BlocksSocket As BlocksSocket)
    If zProtected_MeHaweTheParentFlag Then
        Parent.zProtected_OnDeliveryReporting BytesDelivered, BytesSize, BlocksSocket
    Else
        RaiseEvent OnDeliveryReporting(BytesDelivered, BytesSize, BlocksSocket)
        LastRefresTiks = GetTickCount
    End If
End Sub
