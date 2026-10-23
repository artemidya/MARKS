Private Function GetInvertSegments(Subpaths As VGCore.Subpaths, sx As Double, sy As Double, ex As Double, ey As Double) As RemovableSegment()
    Dim n As Long
    n = Subpaths.count
    Dim rs() As RemovableSegment
    ReDim rs(n)
    
    With rs(0)
        .startX = sx
        .startY = sy
    End With
    
    With rs(n)
        .EndX = ex
        .EndY = ey
    End With
    
    Dim q As Long
    For q = 1 To n
        With Subpaths(q)
            rs(q - 1).EndX = .StartNode.PositionX
            rs(q - 1).EndY = .StartNode.PositionY
            rs(q).startX = .EndNode.PositionX
            rs(q).startY = .EndNode.PositionY
        End With
    Next
    

    If rs(0).startX = rs(0).EndX And rs(0).startY = rs(0).EndY Then
        If UBound(rs) > 0 Then
            For q = 0 To UBound(rs) - 1
                rs(q) = rs(q + 1)
            Next
            ReDim Preserve rs(UBound(rs) - 1)
        Else
            Erase rs()
        End If
    End If

    
    If Not Not rs() Then
        If rs(UBound(rs)).startX = rs(UBound(rs)).EndX And rs(UBound(rs)).startY = rs(UBound(rs)).EndY Then
            If UBound(rs) > 0 Then
                ReDim Preserve rs(UBound(rs) - 1)
            Else
                Erase rs()
            End If
        End If
    End If
    
    GetInvertSegments = rs
End Function
