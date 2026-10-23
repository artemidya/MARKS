Public Property Get ActiveSpeedUP() As Long
    If lActiveSpeedUP > SpeedsCount Then lActiveSpeedUP = SpeedsCount
    If lActiveSpeedUP < 0 Then lActiveSpeedUP = 0
    ActiveSpeedUP = lActiveSpeedUP
End Property
