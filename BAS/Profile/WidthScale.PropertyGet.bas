Public Property Get WidthScale() As Double
    If lWidthScale > 5 Then lWidthScale = 5
    If lWidthScale < -5 Then lWidthScale = -5
    If lWidthScale >= 0 And lWidthScale < 0.5 Then lWidthScale = 0.5
    If lWidthScale > -0.5 And lWidthScale < 0 Then lWidthScale = -0.5
    WidthScale = lWidthScale
End Property
