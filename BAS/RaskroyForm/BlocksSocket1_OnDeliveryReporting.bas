Private Sub BlocksSocket1_OnDeliveryReporting(BytesDelivered As Long, BytesSize As Long, client As BlocksSocket)
    Label1.Caption = localization.ActiveLng.Item("RaskroyForm.SendingDataServer") & "   (" & BytesDelivered \ 1000 & "/" & BytesSize \ 1000 & " KB)"
    ProgressBar1.value = (CDbl(BytesDelivered) / CDbl(BytesSize)) * 10000
End Sub
