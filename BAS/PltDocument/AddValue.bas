Private Function AddValue(value As Long, ByRef SVGtxt As String) As String
    Dim rs As Long
    rs = Asc(right(SVGtxt, 1))
    If rs > 47 And rs < 58 Then
        'need separator
        If value < 0 Then
            AddValue = value
        Else
            AddValue = " " & value
        End If
    Else
        AddValue = value
    End If
End Function
