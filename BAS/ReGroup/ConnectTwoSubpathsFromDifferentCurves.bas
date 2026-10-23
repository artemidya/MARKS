Private Function ConnectTwoSubpathsFromDifferentCurves(SubPath1 As VGCore.SubPath, SubPath2 As VGCore.SubPath) As Boolean
    On Error GoTo 10
    If SubPath1.EndNode.PositionX = SubPath2.StartNode.PositionX Then
        If SubPath1.EndNode.PositionY = SubPath2.StartNode.PositionY Then
            ConnectTwoSubpathsFromDifferentCurves = True
            SubPath1.Parent.AppendCurve SubPath2.Extract.curve
            SubPath2.Delete
            SubPath1.EndNode.ConnectWith SubPath1.Parent.Subpaths.Last.StartNode
            Exit Function
        End If
    End If
'    If SubPath1.EndNode.PositionX = SubPath2.EndNode.PositionX Then
'        If SubPath1.EndNode.PositionY = SubPath2.EndNode.PositionY Then
'            ConnectTwoSubpathsFromDifferentCurves = True
'            SubPath1.Parent.AppendCurve SubPath2.Extract.curve
'            SubPath2.Delete
'            SubPath1.EndNode.ConnectWith SubPath1.Parent.Subpaths.Last.EndNode
'            Exit Function
'        End If
'    End If
'    If SubPath1.StartNode.PositionX = SubPath2.StartNode.PositionX Then
'        If SubPath1.StartNode.PositionY = SubPath2.StartNode.PositionY Then
'             ConnectTwoSubpathsFromDifferentCurves = True
'            SubPath1.Parent.AppendCurve SubPath2.Extract.curve
'            SubPath2.Delete
'            SubPath1.StartNode.ConnectWith SubPath1.Parent.Subpaths.Last.StartNode
'            Exit Function
'        End If
'    End If
    If SubPath1.StartNode.PositionX = SubPath2.EndNode.PositionX Then
        If SubPath1.StartNode.PositionY = SubPath2.EndNode.PositionY Then
            ConnectTwoSubpathsFromDifferentCurves = True
            SubPath1.Parent.AppendCurve SubPath2.Extract.curve
            SubPath2.Delete
            SubPath1.StartNode.ConnectWith SubPath1.Parent.Subpaths.Last.EndNode
            Exit Function
        End If
    End If
  
10
End Function
