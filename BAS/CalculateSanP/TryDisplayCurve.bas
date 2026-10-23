Private Function TryDisplayCurve(ByRef shape As VGCore.shape) As Boolean
    On Error GoTo 10
    LengthOfCurves = LengthOfCurves + Length(shape.DisplayCurve)
    AreaOfCurves = AreaOfCurves + Area(shape.DisplayCurve)
    RectAreaOfCurves = RectAreaOfCurves + RectArea(shape.DisplayCurve)
    
    CountSubPaths = CountSubPaths + shape.DisplayCurve.Subpaths.count
    CountNodes = CountNodes + shape.DisplayCurve.Nodes.count
    
    TryDisplayCurve = True
    Exit Function
10
End Function
