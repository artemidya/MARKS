Private Sub DeleteShapeFromShapesArray(ShapesArray() As VGCore.shape, ByVal q As Long)
    Dim q2 As Long
    For q2 = q To UBound(ShapesArray) - 1
        Set ShapesArray(q2) = ShapesArray(q2 + 1)
    Next
    ReDim Preserve ShapesArray(UBound(ShapesArray) - 1)
End Sub
