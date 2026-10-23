Private Sub SelfyConnectingShapes(sr As VGCore.ShapeRange)
    Dim q As Long
    For q = 1 To sr.count
        If sr(q).Type = cdrCurveShape Then SelfyConnecting sr(q).curve
    Next
End Sub
