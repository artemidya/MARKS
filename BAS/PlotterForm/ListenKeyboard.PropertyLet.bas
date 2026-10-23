Private Property Let ListenKeyboard(oValue As Boolean)
    lListenKeyboard = oValue
    If oValue Then
        'Reset buffer
        Dim q As Long
        For q = 0 To 12 'Count reading keys
            PressedKey
        Next
    Else
        CommandButton23.Picture = Image3.Picture
    End If
End Property
