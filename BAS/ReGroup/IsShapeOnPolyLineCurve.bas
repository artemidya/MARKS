Private Function IsShapeOnPolyLineCurve(crv As VGCore.curve, x As Double, y As Double, shape As VGCore.shape) As Boolean
    With shape
        If .Type = cdrCurveShape Then
            If .curve.closed Then
                If IsOnPolyLineCurve(crv, x, y) Then
                    IsShapeOnPolyLineCurve = True
                Else
                    IsShapeOnPolyLineCurve = IsOnPolyLineCurve(crv, .curve.Nodes.First.PositionX, .curve.Nodes.First.PositionY)
                End If

            Else
                If IsOnPolyLineCurve(crv, .curve.Nodes.First.PositionX, .curve.Nodes.First.PositionY) Then
                    IsShapeOnPolyLineCurve = True
                Else
                    If IsOnPolyLineCurve(crv, .curve.Nodes.Last.PositionX, .curve.Nodes.Last.PositionY) Then
                        IsShapeOnPolyLineCurve = True
                    Else
                        IsShapeOnPolyLineCurve = IsOnPolyLineCurve(crv, x, y)
                    End If
                End If
            End If
        ElseIf .Type = cdrGroupShape Then
            IsShapeOnPolyLineCurve = IsShapeOnPolyLineCurve(crv, x, y, .shapes(1))
        Else
            IsShapeOnPolyLineCurve = IsOnPolyLineCurve(crv, x, y)
        End If
    End With
End Function
