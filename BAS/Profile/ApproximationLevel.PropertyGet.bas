Public Property Get ApproximationLevel() As ApproximationLevels
    If lApproximationLevel > 2 Or lApproximationLevel < 0 Then lApproximationLevel = Normal
    ApproximationLevel = lApproximationLevel
End Property
