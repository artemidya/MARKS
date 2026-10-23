Private Function GetFlowControlListIndex(Combo As MSForms.ComboBox, ComSettingsStr As String) As Long
    Dim parts As Variant
    Dim xon As String, rts As String, dtr As String, odsr As String
    Dim i As Long

    ' –азбиваем строку настроек по пробелам
    parts = Split(ComSettingsStr, " ")

    ' »щем нужные параметры
    For i = LBound(parts) To UBound(parts)
        Select Case LCase(left(parts(i), 4))
            Case "xon="
                xon = LCase(Mid(parts(i), 5))
            Case "rts="
                rts = LCase(Mid(parts(i), 5))
            Case "dtr="
                dtr = LCase(Mid(parts(i), 5))
            Case "odsr"
                odsr = LCase(Mid(parts(i), 6))
        End Select
    Next i

    ' ѕо умолчанию Ч нет управлени€ потоком
    If xon <> "on" And rts <> "handshake" And dtr <> "handshake" And odsr <> "on" Then
        GetFlowControlListIndex = ComboBoxIndexOf(Combo, "None")
    ElseIf xon = "on" Then
        GetFlowControlListIndex = ComboBoxIndexOf(Combo, "Xon/Xoff")
    ElseIf rts = "handshake" Or rts = "on" Then
        GetFlowControlListIndex = ComboBoxIndexOf(Combo, "RTS/CTS")
    ElseIf dtr = "handshake" Or odsr = "on" Then
        GetFlowControlListIndex = ComboBoxIndexOf(Combo, "DSR/DTR")
    Else
        GetFlowControlListIndex = 0 ' не распознано
    End If
End Function
