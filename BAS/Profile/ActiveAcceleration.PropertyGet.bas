Public Property Get ActiveAcceleration() As Long
    If lActiveAcceleration > AccelerationsCount Then lActiveAcceleration = AccelerationsCount
    If lActiveAcceleration < 0 Then lActiveAcceleration = 0
    ActiveAcceleration = lActiveAcceleration
End Property
