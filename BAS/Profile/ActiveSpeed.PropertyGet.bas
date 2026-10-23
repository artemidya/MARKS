Public Property Get ActiveSpeed() As Long
    If lActiveSpeed > SpeedsCount Then lActiveSpeed = SpeedsCount
    If lActiveSpeed < 0 Then lActiveSpeed = 0
    ActiveSpeed = lActiveSpeed
End Property
