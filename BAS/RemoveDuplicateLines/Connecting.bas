Private Sub Connecting(Curves() As VGCore.curve)
    Dim q As Long, q1 As Long, q2 As Long
    For q = 0 To UBound(Curves)
        With Curves(q)
            If .segments.count > 0 Then
                q1 = 1
                Do While q1 <= .Subpaths.count
                    q2 = 1
                    Do While q2 <= .Subpaths.count
                        If q1 <> q2 Then
                            If .Subpaths(q1).EndNode.PositionX = .Subpaths(q2).StartNode.PositionX Then
                                If .Subpaths(q1).EndNode.PositionY = .Subpaths(q2).StartNode.PositionY Then
                                    .Subpaths(q1).EndNode.ConnectWith .Subpaths(q2).StartNode
                                    If q1 > q2 Then
                                        q1 = q1 - 2
                                    Else
                                        q1 = q1 - 1
                                    End If
                                    GoTo ExitLoop
                                End If
                            End If
                        End If
                        q2 = q2 + 1
                    Loop
ExitLoop:
                    q1 = q1 + 1
                Loop
            End If
        End With
    Next
End Sub
