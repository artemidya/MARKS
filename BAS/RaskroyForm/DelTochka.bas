Private Function DelTochka(ByRef txt As String) As Double
    DelTochka = val(Replace$(Replace$(txt, ",", "."), "-", ""))
End Function
