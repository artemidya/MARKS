Private Function KillBirdsAndSpaces(ByVal s As String) As String
    ' Удалить 4 пробела в начале, если строка достаточно длинная
    If Len(s) >= 4 Then
        If left(s, 4) = "    " Then
            KillBirdsAndSpaces = Mid(s, 5)
            Exit Function
        End If
    End If

    ' Удалить символ ChrW(10003) (?) в начале, если строка не пустая
    If Len(s) >= 1 Then
        If left(s, 1) = ChrW(10003) Then
            KillBirdsAndSpaces = Mid(s, 2)
            Exit Function
        End If
    End If

    ' Вернуть как есть, если ничего не найдено
    KillBirdsAndSpaces = s
End Function
