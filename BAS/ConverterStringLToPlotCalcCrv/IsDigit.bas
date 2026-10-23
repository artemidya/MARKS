Private Function IsDigit(ByRef StringL As String) As Boolean
    '[48-57] - Digits "-" - 45
    Static ascii As Long
    ascii = Asc(Mid(StringL, cp, 1))
    If (ascii > 47 And ascii < 58) Or ascii = 45 Then IsDigit = True
End Function
