Property Get Bounds() As RECTL
    If NeedRecalculateBounds Then CalculateBounds
    Bounds = lBounds
End Property
