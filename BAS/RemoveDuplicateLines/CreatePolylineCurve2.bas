Private Function CreatePolylineCurve2(BezierSegment As VGCore.segment, CountPolyLines As Long) As VGCore.curve
    Set CreatePolylineCurve2 = BezierSegment.GetPolyline(CountPolyLines)
End Function
