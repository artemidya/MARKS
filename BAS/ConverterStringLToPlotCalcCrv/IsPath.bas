Private Function IsPath(ByRef StringL As String) As Boolean
    Static bstr As String
    bstr = Mid(StringL, cp, 1)
    If bstr = "U" Or bstr = "D" Or bstr = "B" Then IsPath = True
End Function
