Public Sub ApplyMatrixTransform(matrix As waMatrix)
    GdipTransformPath Ptr, matrix.Ptr
    NeedRecalculateBounds = True
End Sub
