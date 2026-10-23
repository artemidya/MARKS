Private Sub GetCurve(crv As VGCore.curve)
    Dim c As Long:    c = 1000
    With crv
'        Dim path As waPath
        Dim q As Long, q2 As Long
        For q = 1 To .Subpaths.count
            With .Subpaths(q)
'                Set path = New waPath
'                path.StartSubPath
                With .StartNode
'                    path.AddLineSegment .PositionX * c, .PositionY * c
                    OutString = OutString & "U" & CLng(.PositionX * c) & "," & CLng(.PositionY * c)
                End With
                For q2 = 1 To .segments.count
                    With .segments(q2)
                        Select Case .Type
                            Case cdrCurveSegment
'                                path.AddBezierSegment .StartNode.PositionX * c, .StartNode.PositionY * c, .StartingControlPointX * c, .StartingControlPointY * c, .EndingControlPointX * c, .EndingControlPointY * c, .EndNode.PositionX * c, .EndNode.PositionY * c
                                 OutString = OutString & "B" & CLng(.StartingControlPointX * c) & "," & CLng(.StartingControlPointY * c) & "," & CLng(.EndingControlPointX * c) & "," & CLng(.EndingControlPointY * c) & "," & CLng(.EndNode.PositionX * c) & "," & CLng(.EndNode.PositionY * c)
                            Case cdrLineSegment
'                                path.AddLineSegment .EndNode.PositionX * c, .EndNode.PositionY * c
                                 OutString = OutString & "D" & CLng(.EndNode.PositionX * c) & "," & CLng(.EndNode.PositionY * c)
                            Case cdrMixedSegments
                                MsgBox "ERR mixed! (cdrMixedSegments)", vbCritical
                                Exit Sub
                        End Select
                    End With
                Next
'                If .Closed Then path.CloseSubPath
            End With
            
'            Curves.Add path
        Next
    End With
End Sub
