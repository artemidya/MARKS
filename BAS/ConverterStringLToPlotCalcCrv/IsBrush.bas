Private Function IsBrush(ByRef StringL As String) As Boolean
    Static bstr As String
    bstr = Mid(StringL, cp, 1)
    If bstr = "S" Then IsBrush = True: Exit Function
    If bstr = "G" Then IsBrush = True
End Function
