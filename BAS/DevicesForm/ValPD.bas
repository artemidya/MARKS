Private Function ValPD(ByRef txt As String) As Double
    ValPD = val(Replace$(Replace$(txt, "%", ""), ",", "."))
End Function
