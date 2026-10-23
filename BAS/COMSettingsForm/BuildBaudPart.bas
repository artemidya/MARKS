Private Function BuildBaudPart(Combo As MSForms.ComboBox) As String
    If Combo.ListIndex >= 0 Then
        BuildBaudPart = "baud=" & Combo.value
    Else
        BuildBaudPart = ""
    End If
End Function
