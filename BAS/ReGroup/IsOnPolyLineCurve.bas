Private Function IsOnPolyLineCurve(crv As VGCore.curve, x As Double, y As Double) As Boolean
    Dim q As Long, q2 As Long
    Dim x1 As Double, y1 As Double
    Dim x2 As Double, y2 As Double
    Dim dy As Double
    
    With crv
        For q = 1 To .Subpaths.count
            With .Subpaths(q)
                If .closed Then
                    If x > .boundingBox.left And x < .boundingBox.right Then
                        If y > .boundingBox.bottom And y < .boundingBox.top Then
                            For q2 = 1 To .segments.count
                                GetX1Y1X2Y2 x1, y1, x2, y2, .segments(q2)
                                If y > y1 Then
                                    If y <= y2 Then
                                        With .segments(q2)
                                            If x > .StartNode.PositionX Or x > .EndNode.PositionX Then
                                                dy = (y2 - y1)
                                                If dy > 0 Then
                                                    If (y - y1) * (x2 - x1) / dy + x1 < x Then
                                                        IsOnPolyLineCurve = Not IsOnPolyLineCurve
                                                    End If
                                                End If
                                            End If
                                        End With
                                    End If
                                End If
                            Next
                        End If
                    End If
                End If
            End With
        Next
    End With
End Function
