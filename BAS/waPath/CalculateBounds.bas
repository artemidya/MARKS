Private Sub CalculateBounds()
    GdipGetPathWorldBoundsI Ptr, lBounds, 0, 0
    With lBounds
        .right = .left + .right - 1
        .bottom = .top + .bottom - 1
    End With
    NeedRecalculateBounds = False
End Sub
