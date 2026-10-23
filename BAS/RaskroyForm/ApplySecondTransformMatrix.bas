Private Sub ApplySecondTransformMatrix(sr As ShapeRange, ByRef mfc() As MatrixForCheck, ByRef transformMatrix As transformMatrix, Optional ByRef qq As Long = -1)
    
    Dim q As Long

    For q = 1 To sr.shapes.count
        qq = qq + 1
        'Compare
        If CompareMatrixes(mfc(qq), sr.shapes(q)) Then
            'Apply
            sr.shapes(q).ApplyTransformMatrix transformMatrix
        End If
        
        If IsPowerClip(sr.shapes(q)) Then
            ApplySecondTransformMatrix sr.shapes(q).PowerClip.shapes.All, mfc(), transformMatrix, qq
        ElseIf sr.shapes(q).Type = cdrGroupShape Then
            ApplySecondTransformMatrix sr.shapes(q).shapes.All, mfc(), transformMatrix, qq
        End If
    Next
End Sub
