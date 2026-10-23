Private Function TryDisplayCurve(shape As VGCore.shape) As Boolean
    On Error GoTo 10
    If shape.DisplayCurve.Nodes.count > 0 Then TryDisplayCurve = True
10
End Function
