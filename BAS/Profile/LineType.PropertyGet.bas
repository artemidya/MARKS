Public Property Get LineType() As LineTypes
    If lLineType > 1 Or lLineType < 0 Then lLineType = lines
    LineType = lLineType
End Property
