Public Property Get Turn() As Turns
    If lTurn > 3 Then lTurn = Turn0
    If lTurn < 0 Then lTurn = Turn0
    Turn = lTurn
End Property
