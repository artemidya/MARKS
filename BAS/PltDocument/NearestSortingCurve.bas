Private Function NearestSortingCurve(CurvesForSearching As Collection, LastPoint As POINTL) As SortingCurve
    Dim CurrentSortingCurve As SortingCurve
    Dim minDistance As Long: minDistance = 2147483647

    Dim minQ As Long
    Dim minNodeNumber As Long
    Dim pp As NearestPointInfo
    Dim d As Long

    Dim q As Long
    Dim lx As Long, ly As Long
    
    With LastPoint
        lx = .x
        ly = .y
    End With
    For q = 1 To CurvesForSearching.count
        Set CurrentSortingCurve = CurvesForSearching(q)
        With CurrentSortingCurve
            If .path.closed Then
                pp = FindNearestPoint(.path, lx, ly)
                With pp
                    If .Lenth < minDistance Then
                        minQ = q
                        minDistance = .Lenth
                        minNodeNumber = .Number
                    End If
                End With
            Else
                With .path.FirstPoint
                    d = Distance(.x, .y, lx, ly)
                End With
                If d < minDistance Then
                    minQ = q
                    minDistance = d
                    minNodeNumber = 0
                End If
                
                With .path.LastPoint
                    d = Distance(.x, .y, lx, ly)
                End With
                If d < minDistance Then
                    minQ = q
                    minDistance = d
                    minNodeNumber = -1
                End If
                
            End If
        End With
    Next
    
    '===
    Set NearestSortingCurve = CurvesForSearching(minQ)
    With NearestSortingCurve
        If .Childs.count = 0 Then
            If minNodeNumber > 0 Then
                Set .path = .path.ChangeStartNode(minNodeNumber)
            ElseIf minNodeNumber < 0 Then
                'Need revers not closed curve
                .path.Reverse
            End If
        End If
    End With
    
    CurvesForSearching.Remove minQ
End Function
