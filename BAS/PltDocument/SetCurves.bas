Private Function SetCurves(SortingCurves As Collection) As Collection
    Dim q As Long
    Dim sc As SortingCurve
    Dim c2 As Collection
    Dim NewCurves As New Collection
    
    With SortingCurves
        For q = 1 To .count
            Set sc = SortingCurves(q)
            With sc
                If .Childs.count > 0 Then
                    Set c2 = SetCurves(.Childs)
                    AddCollectionToCollection NewCurves, c2
                End If
                NewCurves.Add .path
            End With
        Next
    End With
    
    Set SetCurves = NewCurves
End Function
