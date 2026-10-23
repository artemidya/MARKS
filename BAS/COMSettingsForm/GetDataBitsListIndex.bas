Private Function GetDataBitsListIndex(Combo As MSForms.ComboBox, ComSettingsStr As String) As Long
    Dim dataBitsStr As String
    Dim dataBits As Long
    Dim parts As Variant
    Dim i As Long

    ' Разбиваем строку настроек по пробелам
    parts = Split(ComSettingsStr, " ")

    ' Ищем параметр data=
    For i = LBound(parts) To UBound(parts)
        If LCase(left(parts(i), 5)) = "data=" Then
            dataBitsStr = Mid(parts(i), 6)
            Exit For
        End If
    Next i

    If dataBitsStr = "" Then
        GetDataBitsListIndex = 3 ' параметр data= не найден
        Exit Function
    End If

    On Error Resume Next
    dataBits = CLng(dataBitsStr)
    On Error GoTo 0

    ' Сравниваем с элементами ComboBox
    For i = 0 To Combo.ListCount - 1
        If CLng(Combo.List(i)) = dataBits Then
            GetDataBitsListIndex = i
            Exit Function
        End If
    Next i

    GetDataBitsListIndex = 3 ' не найдено
End Function
