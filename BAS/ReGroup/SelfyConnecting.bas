Private Sub SelfyConnecting(crv As VGCore.curve)
    Dim q As Long, q2 As Long
    With crv
        q = 1
        Do While q <= .Subpaths.count
            If Not .Subpaths(q).closed Then
                q2 = q + 1
                Do While q2 <= .Subpaths.count
                    If Not .Subpaths(q2).closed Then
                        If ConnectTwoSubpaths(.Subpaths(q), .Subpaths(q2)) Then
                            q2 = q + 1
                            GoTo nx:
                        End If
                    End If
                    q2 = q2 + 1
nx:
                Loop
                If .Subpaths(q).Nodes.Last.PositionX = .Subpaths(q).Nodes.First.PositionX Then
                    If .Subpaths(q).Nodes.Last.PositionY = .Subpaths(q).Nodes.First.PositionY Then
                        If .Subpaths(q).Nodes.count > 1 Then
                            .Subpaths(q).closed = True
                        End If
                    End If
                End If
            End If
            q = q + 1
        Loop
    End With
End Sub
