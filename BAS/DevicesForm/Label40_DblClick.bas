Private Sub Label40_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    If Not Connected Then
        BlocksSocket1.Connect UserData.MainAddress, UserData.MainAddressPort
    End If
End Sub
