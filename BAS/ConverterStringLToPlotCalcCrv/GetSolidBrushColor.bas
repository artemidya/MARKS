Private Function GetSolidBrushColor(ByRef StringL As String) As RGBA
    Dim d As String
    GetSolidBrushColor.a = 255
    
    cp = cp + 1
    Do While cp <= LenStringL
        If IsDigit(StringL) Then d = d & Mid(StringL, cp, 1) Else Exit Do
        cp = cp + 1
    Loop
    GetSolidBrushColor.R = TxtToByte(d)
    
    cp = cp + 1
    d = ""
    Do While cp <= LenStringL
        If IsDigit(StringL) Then d = d & Mid(StringL, cp, 1) Else Exit Do
        cp = cp + 1
    Loop
    GetSolidBrushColor.g = TxtToByte(d)
    
    cp = cp + 1
    d = ""
    Do While cp <= LenStringL
        If IsDigit(StringL) Then d = d & Mid(StringL, cp, 1) Else Exit Do
        cp = cp + 1
    Loop
    GetSolidBrushColor.b = TxtToByte(d)
    
    cp = cp + 1
    d = ""
    Do While cp <= LenStringL
        If IsDigit(StringL) Then d = d & Mid(StringL, cp, 1) Else Exit Do
        cp = cp + 1
    Loop
    GetSolidBrushColor.a = TxtToByte(d)
    
End Function
