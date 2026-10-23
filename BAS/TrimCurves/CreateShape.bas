Private Function CreateShape(crv As VGCore.curve) As VGCore.shape
    Set CreateShape = VGCore.ActiveLayer.CreateCurve(crv)
    Dim q As Long
    With CreateShape.curve
        For q = 1 To .Subpaths.count
            ClosePathIfDistanceLessThan .Subpaths(q), BreakDistance * 0.9
            If .Subpaths(q).closed = True And .Subpaths(q).Length > BreakDistance * 1.1 Then
                MaxNodeInSubPath(.Subpaths(q)).BreakApart
                TrimSubPath .Subpaths(q), BreakDistance / 2
            End If
        Next
    End With
End Function
