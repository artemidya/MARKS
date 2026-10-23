Public Function ChangeStartNode(NumberOfNewStartNode As Long) As waPath
    If NumberOfNewStartNode < 2 Then
        Set ChangeStartNode = Me
        Exit Function
    End If
    
    Dim segments() As segment
    segments = Me.segments
    
    Dim path As waPath
    Set path = New waPath
    path.StartSubPath
    With segments(NumberOfNewStartNode).P1
        path.AddLineSegment .x, .y
    End With
    Dim q As Long
    For q = NumberOfNewStartNode To UBound(segments)
        With segments(q)
            If .Type = CurveSegment Then
                path.AddBezierSegment .P1.x, .P1.y, .P2.x, .P2.y, .P3.x, .P3.y, .P4.x, .P4.y
            ElseIf .Type = LineSegment Then
                path.AddLineSegment .P2.x, .P2.y
            End If
        End With
    Next
    
    For q = 1 To NumberOfNewStartNode - 1
        With segments(q)
            If .Type = CurveSegment Then
                path.AddBezierSegment .P1.x, .P1.y, .P2.x, .P2.y, .P3.x, .P3.y, .P4.x, .P4.y
            ElseIf .Type = LineSegment Then
                path.AddLineSegment .P2.x, .P2.y
            End If
        End With
    Next
    Set path.brush = Me.brush
    Set path.pen = Me.pen
    Set ChangeStartNode = path
    
End Function
