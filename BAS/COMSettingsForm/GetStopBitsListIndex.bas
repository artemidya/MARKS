Private Function GetStopBitsListIndex(Combo As MSForms.ComboBox, ComSettingsStr As String) As Long
    Dim stopBitsStr As String
    Dim parts As Variant
    Dim i As Long

    ' Разбиваем строку на отдельные параметры
    parts = Split(ComSettingsStr, " ")

    ' Ищем параметр stop=
    For i = LBound(parts) To UBound(parts)
        If LCase(left(parts(i), 5)) = "stop=" Then
            stopBitsStr = Mid(parts(i), 6)
            Exit For
        End If
    Next i

    If stopBitsStr = "" Then
        GetStopBitsListIndex = 0 ' stop не найден
        Exit Function
    End If

    ' Нормализуем: "1.5" остаётся строкой, "2" — строка и т.д.
    stopBitsStr = Trim(stopBitsStr)

    ' Ищем в ComboBox
    For i = 0 To Combo.ListCount - 1
        If Combo.List(i) = stopBitsStr Then
            GetStopBitsListIndex = i
            Exit Function
        End If
    Next i

    GetStopBitsListIndex = 0 ' не найдено
End Function
