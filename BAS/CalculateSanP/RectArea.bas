Private Function RectArea(ByRef crv As VGCore.curve) As Double
    Dim px() As Double, Py() As Double
    Dim rX() As Double, dy() As Double
    Dim deleted() As Boolean
    ReDim px(crv.Subpaths.count), Py(crv.Subpaths.count)
    ReDim rX(crv.Subpaths.count), dy(crv.Subpaths.count)
    ReDim deleted(crv.Subpaths.count)
    Dim q As Long, q2 As Long
    
    For q = 1 To crv.Subpaths.count
        px(q) = crv.Subpaths(q).boundingBox.left
    Next
    
    For q = 1 To crv.Subpaths.count
        Py(q) = crv.Subpaths(q).boundingBox.top
    Next
    
    For q = 1 To crv.Subpaths.count
        rX(q) = crv.Subpaths(q).boundingBox.right
    Next
    
    For q = 1 To crv.Subpaths.count
        dy(q) = crv.Subpaths(q).boundingBox.bottom
    Next

    For q = 1 To crv.Subpaths.count
        For q2 = 1 To crv.Subpaths.count
            If q <> q2 Then
                If px(q) > px(q2) Then
                    If rX(q) < rX(q2) Then
                        If Py(q) < Py(q2) Then
                            If dy(q) > dy(q2) Then
                                deleted(q) = True
                                Exit For
                            End If
                        End If
                    End If
                End If
            End If
        Next
    Next
    
    For q = 1 To crv.Subpaths.count
        If deleted(q) = False Then
            RectArea = RectArea + (rX(q) - px(q)) * (Py(q) - dy(q))
            RectOnScreenCurvesCount = RectOnScreenCurvesCount + 1
            ReDim Preserve RectOnScreenCurves(RectOnScreenCurvesCount)
            Set RectOnScreenCurves(RectOnScreenCurvesCount) = New OnScreenCurve
            RectOnScreenCurves(RectOnScreenCurvesCount).SetRectangle px(q), Py(q), rX(q), dy(q)
            RectOnScreenCurves(RectOnScreenCurvesCount).SetBrush RGB(230, 230, 255)
        End If
    Next
End Function
