Private Function ValP(str As String) As Double
    str = Replace(str, ",", ".")
    ValP = val(str)
End Function
