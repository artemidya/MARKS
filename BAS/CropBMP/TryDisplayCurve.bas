Private Function TryDisplayCurve(shape As shape) As Boolean
    On Error GoTo Err
    Dim c As Long
    c = shape.DisplayCurve.Subpaths.count
    TryDisplayCurve = True
    Exit Function
Err:
End Function
