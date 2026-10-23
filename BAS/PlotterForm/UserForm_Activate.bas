Private Sub UserForm_Activate()
    If Not UserForm_Activ Then
        UserForm_Activ = True
        'Read data from CorelDraw
        Set OriginalCurves = Converter.GetCurves(OriginalHeight, OriginalWidth, MaxColor)
        
        'Init BlocksSokets
        BlocksSocketShared.HookForm FindWindow(vbNullString, Me.Caption)
        Set BlocksSocket1 = New BlocksSocket
        BlocksSocket1.Connect UserData.MainAddress, UserData.MainAddressPort
        Set USB_Events1 = New USB_Events
        
        
        If OriginalCurves.count = 0 Then Command5.Enabled = False
        Label21.Visible = False
        If OriginalCurves.count > 0 Then CommandButton24.Enabled = True
        RefreshPltDocument
        
        KeyReadLoop
    End If
End Sub
