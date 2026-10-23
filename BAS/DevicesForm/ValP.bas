Private Function ValP(ByRef txt As String) As Double
    ValP = Abs(val(Replace$(Replace$(txt, "%", ""), ",", ".")))
End Function
