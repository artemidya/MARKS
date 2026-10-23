Private Function GetBaudListIndex(Combo As MSForms.ComboBox, ComSettingsStr As String) As Long
    Dim baudRateStr As String
    Dim baudRate As Long
    Dim parts As Variant
    Dim i As Long

    ' Разбираем строку на части
    parts = Split(ComSettingsStr, " ")
    
    ' Ищем часть, начинающуюся с "baud="
    For i = LBound(parts) To UBound(parts)
        If LCase(left(parts(i), 5)) = "baud=" Then
            baudRateStr = Mid(parts(i), 6) ' пропускаем "baud="
            Exit For
        End If
    Next i

    If baudRateStr = "" Then
        GetBaudListIndex = 5 ' baud не найден
        Exit Function
    End If

    ' Преобразуем в число
    On Error Resume Next
    baudRate = CLng(baudRateStr)
    On Error GoTo 0

    ' Ищем индекс в ComboBox
    For i = 0 To Combo.ListCount - 1
        If CLng(Combo.List(i)) = baudRate Then
            GetBaudListIndex = i
            Exit Function
        End If
    Next i

    ' Если не найдено — вернуть -5
    GetBaudListIndex = 5
End Function
