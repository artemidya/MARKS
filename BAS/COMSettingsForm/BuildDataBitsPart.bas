Private Function BuildDataBitsPart(Combo As MSForms.ComboBox) As String
    If Combo.ListIndex >= 0 Then
        BuildDataBitsPart = "data=" & Combo.value
    Else
        BuildDataBitsPart = ""
    End If
End Function
