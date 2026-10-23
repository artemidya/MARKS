Private Function GetMatrixForCheck(sr As ShapeRange) As MatrixForCheck()
    Dim mfc() As MatrixForCheck
    Dim q As Long

    For q = 1 To sr.shapes.count
        If Not Not mfc() Then
            ReDim Preserve mfc(UBound(mfc()) + 1)
        Else
            ReDim mfc(0)
        End If
        
        With sr.shapes(q).TransformationMatrix
            mfc(UBound(mfc)).d11 = .d11
            mfc(UBound(mfc)).d12 = .d12
            mfc(UBound(mfc)).d21 = .d21
            mfc(UBound(mfc)).d22 = .d22
        End With
        
        If IsPowerClip(sr.shapes(q)) Then
            AddArrayToArray mfc, GetMatrixForCheck(sr.shapes(q).PowerClip.shapes.All)
        ElseIf sr.shapes(q).Type = cdrGroupShape Then
            AddArrayToArray mfc, GetMatrixForCheck(sr.shapes(q).shapes.All)
        End If
    Next
    GetMatrixForCheck = mfc()
End Function
