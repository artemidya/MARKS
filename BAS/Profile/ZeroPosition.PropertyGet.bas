Public Property Get ZeroPosition() As ZeroPositions
    If lZeroPosition > 2 Or lZeroPosition < 0 Then lZeroPosition = RightP
    ZeroPosition = lZeroPosition
End Property
