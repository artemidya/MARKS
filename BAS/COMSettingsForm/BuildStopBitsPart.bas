Private Function BuildStopBitsPart(Combo As MSForms.ComboBox) As String
    If Combo.ListIndex >= 0 Then
        BuildStopBitsPart = "stop=" & Combo.value
    Else
        BuildStopBitsPart = ""
    End If
End Function
