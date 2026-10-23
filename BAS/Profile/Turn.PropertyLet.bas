Public Property Let Turn(oValue As Turns)
    If oValue <> lTurn Then NeedSave = True
    lTurn = oValue
End Property
