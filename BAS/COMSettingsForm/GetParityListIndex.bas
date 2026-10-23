Private Function GetParityListIndex(Combo As MSForms.ComboBox, ComSettingsStr As String) As Long
    Dim parityStr As String
    Dim parts As Variant
    Dim i As Long

    ' Разбиваем строку настроек по пробелам
    parts = Split(ComSettingsStr, " ")

    ' Ищем параметр parity=
    For i = LBound(parts) To UBound(parts)
        If LCase(left(parts(i), 7)) = "parity=" Then
            parityStr = Mid(parts(i), 8)
            Exit For
        End If
    Next i

    If parityStr = "" Then
        GetParityListIndex = 0 ' parity не найден
        Exit Function
    End If

    ' Приводим к удобному формату для сравнения
    Select Case UCase(parityStr)
        Case "N": parityStr = "None"
        Case "O": parityStr = "Odd"
        Case "E": parityStr = "Even"
        Case "M": parityStr = "Mark"
        Case "S": parityStr = "Space"
        ' другие — игнорируем
    End Select

    ' Ищем соответствие в ComboBox
    For i = 0 To Combo.ListCount - 1
        If StrComp(Combo.List(i), parityStr, vbTextCompare) = 0 Then
            GetParityListIndex = i
            Exit Function
        End If
    Next i

    GetParityListIndex = 0 ' не найдено
End Function
