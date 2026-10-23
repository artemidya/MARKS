Private Function ValP(ByRef txt As String) As Double
    ValP = Abs(val(Replace$(txt, ",", ".")))
End Function
