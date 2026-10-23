Private Function CreatePolylineCurveFromCurve(crv As VGCore.curve) As VGCore.curve
    Set CreatePolylineCurveFromCurve = Application.CreateCurve
    Dim q As Long, q2 As Long, q3 As Long
    Dim sp As VGCore.SubPath
    Dim plc As VGCore.curve
    
    With crv
        For q = 1 To .Subpaths.count
            With .Subpaths(q)
                Set sp = CreatePolylineCurveFromCurve.CreateSubPath(.StartNode.PositionX, .StartNode.PositionY)
                
                For q2 = 1 To .segments.count
                    With .segments(q2)
                        If .Type = cdrCurveSegment Then
                            Set plc = .GetPolyline
                            With plc
                                For q3 = 1 To plc.segments.count
                                    With .segments(q3)
                                        sp.AppendLineSegment .EndNode.PositionX, .EndNode.PositionY
                                    End With
                                Next
                            End With
                        Else
                            sp.AppendLineSegment .EndNode.PositionX, .EndNode.PositionY
                        End If
                    End With
                Next
                sp.closed = .closed
            End With
        Next
    End With
End Function
