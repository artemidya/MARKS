Private Function ConnectTwoSubpaths(SubPath1 As VGCore.SubPath, SubPath2 As VGCore.SubPath) As Boolean
    If SubPath1.EndNode.PositionX = SubPath2.StartNode.PositionX Then
        If SubPath1.EndNode.PositionY = SubPath2.StartNode.PositionY Then
            SubPath1.EndNode.ConnectWith SubPath2.StartNode
            ConnectTwoSubpaths = True
            Exit Function
        End If
    End If
    If SubPath1.EndNode.PositionX = SubPath2.EndNode.PositionX Then
        If SubPath1.EndNode.PositionY = SubPath2.EndNode.PositionY Then
            SubPath1.EndNode.ConnectWith SubPath2.EndNode
            ConnectTwoSubpaths = True
            Exit Function
        End If
    End If
    If SubPath1.StartNode.PositionX = SubPath2.StartNode.PositionX Then
        If SubPath1.StartNode.PositionY = SubPath2.StartNode.PositionY Then
            SubPath1.StartNode.ConnectWith SubPath2.StartNode
            ConnectTwoSubpaths = True
            Exit Function
        End If
    End If
    If SubPath1.StartNode.PositionX = SubPath2.EndNode.PositionX Then
        If SubPath1.StartNode.PositionY = SubPath2.EndNode.PositionY Then
            SubPath1.StartNode.ConnectWith SubPath2.EndNode
            ConnectTwoSubpaths = True
            Exit Function
        End If
    End If
End Function
