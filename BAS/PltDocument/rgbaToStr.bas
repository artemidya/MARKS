Private Function rgbaToStr(Color As RGBA) As String
    Dim s As Double
    Dim aa As String
    
    With Color
        s = .a / 255
        aa = Replace(Format(s, "#0.00"), ",", ".")
        If aa = "1.00" Then aa = 1
        rgbaToStr = "rgba(" & .R & "," & .g & "," & .b & "," & aa & ")"
    End With
End Function
