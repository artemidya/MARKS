Private Function BuildFlowControlParts(Combo As MSForms.ComboBox) As String
    If Combo.ListIndex < 0 Then
        BuildFlowControlParts = ""
        Exit Function
    End If

    Select Case UCase(Combo.value)
        Case "NONE"
            BuildFlowControlParts = "xon=off odsr=off dtr=on rts=on"
        Case "XON/XOFF"
            BuildFlowControlParts = "xon=on odsr=off dtr=on rts=on"
        Case "RTS/CTS"
            BuildFlowControlParts = "xon=off odsr=off dtr=on rts=handshake"
        Case "DSR/DTR"
            BuildFlowControlParts = "xon=off odsr=on dtr=handshake rts=on"
        Case Else
            BuildFlowControlParts = "xon=off odsr=off dtr=on rts=on"
    End Select
End Function
