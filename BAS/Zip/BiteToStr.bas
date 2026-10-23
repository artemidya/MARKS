Private Function BiteToStr(ByRef m() As Byte) As String
Dim q As Long
For q = 0 To UBound(m)
    BiteToStr = BiteToStr & Chr(m(q))
Next
End Function
