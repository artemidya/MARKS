Private Function ComboBoxIndexOf(Combo As MSForms.ComboBox, searchText As String) As Long
    Dim i As Long
    For i = 0 To Combo.ListCount - 1
        If StrComp(Combo.List(i), searchText, vbTextCompare) = 0 Then
            ComboBoxIndexOf = i
            Exit Function
        End If
    Next i
    ComboBoxIndexOf = 0
End Function
