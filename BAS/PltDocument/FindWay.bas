Private Function FindWay(SortingCurves As Collection, PreviosPoint As POINTL) As Collection
    Dim NewCurves As New Collection
    Dim sc As SortingCurve
    
    Dim lsc As SortingCurve
    Dim pp As NearestPointInfo
    Dim d1 As Long, d2 As Long
    Dim LastPoint As POINTL
    Dim OldHorisontalDelitel As Long
    
    LastPoint = PreviosPoint
    
    
    With SortingCurves
        Do While .count > 0
            Set sc = NearestSortingCurve(SortingCurves, LastPoint)
            With sc
                If .Childs.count > 0 Then
                    Set .Childs = FindWay(.Childs, LastPoint)
                    
                    Set lsc = .Childs(.Childs.count)
                    LastPoint = lsc.path.LastPoint
                        
                    If .path.closed Then
                        OldHorisontalDelitel = HorisontalDelitel
                        HorisontalDelitel = 1
                        pp = FindNearestPoint(.path, LastPoint.x * HorisontalDelitel, LastPoint.y)
                        HorisontalDelitel = OldHorisontalDelitel
                        Set .path = .path.ChangeStartNode(pp.Number)
                    Else
                        With .path.FirstPoint
                            d1 = Distance(.x, .y, LastPoint.x, LastPoint.y)
                        End With
                        With .path.LastPoint
                            d2 = Distance(.x, .y, LastPoint.x, LastPoint.y)
                        End With
                        If d1 > d2 Then .path.Reverse
                    End If
                End If
                LastPoint = .path.LastPoint
                NewCurves.Add sc
                
            End With
        Loop
    End With
    '====FINAL=========================================================================
    Dim q As Long
    For q = 2 To NewCurves.count
        FinalFinding NewCurves, q, PreviosPoint
    Next
    '==================================================================================
    Set FindWay = NewCurves
End Function
