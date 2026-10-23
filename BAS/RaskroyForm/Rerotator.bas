Private Sub Rerotator(ByRef shiftX As Double, ByRef shiftY As Double, crv As curve, sr As ShapeRange, Angel As Double)
    Dim transformMatrix As New transformMatrix
    transformMatrix.Rotate Angel
      
    Dim mfc() As MatrixForCheck
    mfc = GetMatrixForCheck(sr)
    
    Dim q As Long

    
    For q = 1 To sr.shapes.count
        sr.shapes(q).ApplyTransformMatrix transformMatrix
    Next
    
    Dim mfc2() As MatrixForCheck
    mfc2 = GetMatrixForCheck(sr)
    

    ApplySecondTransformMatrix sr, mfc, transformMatrix
    

    crv.ApplyTransformMatrix transformMatrix
        
    shiftX = crv.boundingBox.right - sr.RightX
    shiftY = crv.boundingBox.top - sr.PositionY
    
End Sub
