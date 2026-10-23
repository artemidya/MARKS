Private Sub BlocksSocket1_OnConnectionProcessIsStarted()
    Label40.Caption = localization.ActiveLng.Item("DevicesForm.Label40.Caption_conTo") & " " & UserData.MainAddress & ":" & UserData.MainAddressPort
End Sub
