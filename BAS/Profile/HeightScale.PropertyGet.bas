Public Property Get HeightScale() As Double
    If lHeightScale > 5 Then lHeightScale = 5
    If lHeightScale < -5 Then lHeightScale = -5
    If lHeightScale >= 0 And lHeightScale < 0.5 Then lHeightScale = 0.5
    If lHeightScale > -0.5 And lHeightScale < 0 Then lHeightScale = -0.5
    HeightScale = lHeightScale
End Property
