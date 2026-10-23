Private Sub CalculateShapes(ByRef shapes As VGCore.shapes)
    Dim q1 As Long
    For q1 = 1 To shapes.count
        If shapes(q1).Type = cdrCurveShape Then
            LengthOfCurves = LengthOfCurves + Length(shapes(q1).curve)
            AreaOfCurves = AreaOfCurves + Area(shapes(q1).curve)
            RectAreaOfCurves = RectAreaOfCurves + RectArea(shapes(q1).curve)
            CountSubPaths = CountSubPaths + shapes(q1).curve.Subpaths.count
            CountNodes = CountNodes + shapes(q1).curve.Nodes.count
            
        ElseIf shapes(q1).Type = cdrGroupShape _
        Or shapes(q1).Type = cdrContourGroupShape _
        Or shapes(q1).Type = cdrArtisticMediaGroupShape Then
             CalculateShapes shapes(q1).shapes
        ElseIf shapes(q1).Type = cdrBitmapShape Then
            CountErrPlus1 shapes(q1)
        Else
            '.......................
            If TryDisplayCurve(shapes(q1)) = False Then
                CountErrPlus1 shapes(q1)
            End If
        End If
        If Timer - t > 0.02 Then
            t = Timer
            If PerimetrFormVisible = True Then PerimetrForm.UserFormRefreshData
            If SquareFormVisible = True Then
                SquareForm.UserFormRefreshData
                If SquareForm.ComboBox1.ListIndex = 1 Then ShowRectOnScreenCurves Else HideRectOnScreenCurves
            End If
            DoEvents
        End If
        If NeedBreak = True Then Exit Sub
    Next
End Sub
