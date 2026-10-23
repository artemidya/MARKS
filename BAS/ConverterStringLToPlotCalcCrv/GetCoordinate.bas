Private Function GetCoordinate(ByRef StringL As String) As Long
    Dim d As String
    
    cp = cp + 1
    Do While cp <= LenStringL
        If IsDigit(StringL) Then d = d & Mid(StringL, cp, 1) Else Exit Do
        cp = cp + 1
    Loop
    GetCoordinate = TxtToLong(d)
End Function
