Private Function CreatePolylineCurveFromShape(shape As VGCore.shape) As VGCore.curve
    With shape
        If .Type = cdrCurveShape Then
            Set CreatePolylineCurveFromShape = CreatePolylineCurveFromCurve(shape.curve)
        ElseIf .Type = cdrEllipseShape Then
            Set CreatePolylineCurveFromShape = CreatePolylineCurveFromCurve(shape.DisplayCurve)
        Else
            Set CreatePolylineCurveFromShape = Application.CreateCurve
            CreatePolylineCurveFromShape.AppendSubpathRectangle .LeftX, .TopY, .RightX, .BottomY
        End If
    End With
End Function
