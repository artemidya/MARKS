Private Sub GetCurve(crv As VGCore.curve, forBrushString As String)
    Dim c As Long:    c = 1000
    With crv
                
        TextBox1 = TextBox1 & vbNewLine & vbTab & "Set path = New waPath" & vbNewLine
        TextBox1 = TextBox1 & vbTab & "With path" & vbNewLine
        Dim q As Long, q2 As Long
        For q = 1 To .Subpaths.count
            With .Subpaths(q)

                TextBox1 = TextBox1 & vbTab & vbTab & ".StartSubPath" & vbNewLine
                With .StartNode
                    TextBox1 = TextBox1 & vbTab & vbTab & ".AddLineSegment " & CLng(.PositionX * c) & ", " & CLng(.PositionY * c) & vbNewLine
                End With
                For q2 = 1 To .segments.count
                    With .segments(q2)
                        Select Case .Type
                            Case cdrCurveSegment
                                TextBox1 = TextBox1 & vbTab & vbTab & ".AddBezierSegment " & CLng(.StartNode.PositionX * c) & ", " & CLng(.StartNode.PositionY * c) & ", " & CLng(.StartingControlPointX * c) & ", " & CLng(.StartingControlPointY * c) & ", " & CLng(.EndingControlPointX * c) & ", " & CLng(.EndingControlPointY * c) & ", " & CLng(.EndNode.PositionX * c) & ", " & CLng(.EndNode.PositionY * c) & vbNewLine
                            Case cdrLineSegment
                                TextBox1 = TextBox1 & vbTab & vbTab & ".AddLineSegment " & CLng(.EndNode.PositionX * c) & ", " & CLng(.EndNode.PositionY * c) & vbNewLine
                        End Select
                    End With
                Next
                If .closed Then TextBox1 = TextBox1 & vbTab & vbTab & ".CloseSubPath" & vbNewLine & vbNewLine

            End With
        Next
        TextBox1 = TextBox1 & forBrushString & vbNewLine
        TextBox1 = TextBox1 & vbTab & "End With" & vbNewLine
        TextBox1 = TextBox1 & vbTab & "myCurves.Add path" & vbNewLine
    End With
End Sub
