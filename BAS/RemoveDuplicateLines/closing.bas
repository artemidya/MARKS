Private Sub closing(Curves() As VGCore.curve)
    Dim q As Long, q2 As Long
    For q = 0 To UBound(Curves)
        With Curves(q)
            If .segments.count > 0 Then
                For q2 = 1 To .Subpaths.count
                    With .Subpaths(q2)
                        If .StartNode.PositionX = .EndNode.PositionX Then
                            If .StartNode.PositionY = .EndNode.PositionY Then
                                .closed = True
                            End If
                        End If
                    End With
                Next
            End If
        End With
    Next
End Sub
