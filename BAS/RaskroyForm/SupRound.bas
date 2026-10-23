Private Function SupRound(ByVal Number As Double) As Double
    'Убирает 4 знак после запятой
    SupRound = Int(Number * 10000) / 10000
End Function
