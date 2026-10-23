Private Function MaxNodeInSubPath(SubPath As VGCore.SubPath) As VGCore.Node
    With SubPath
        If .Nodes.count > 1 Then
            Dim maxY As Double
            Dim q As Long
            Set MaxNodeInSubPath = .Nodes(1)
            With .segments(1)
                If .Type = cdrLineSegment Then
                    maxY = (.StartNode.PositionY + .EndNode.PositionY) / 2
                Else
                    maxY = .StartNode.PositionY
                End If
            End With
            For q = 2 To .segments.count
                With .segments(q)
                    If .Type = cdrLineSegment Then
                        If (.StartNode.PositionY + .EndNode.PositionY) / 2 > maxY Then
                            maxY = (.StartNode.PositionY + .EndNode.PositionY) / 2
                            Set MaxNodeInSubPath = .StartNode
                        End If
                    Else
                        If .StartNode.PositionY > maxY Then
                            maxY = .StartNode.PositionY
                            Set MaxNodeInSubPath = .StartNode
                        End If
                    End If
                End With
            Next
        End If
    End With
End Function
