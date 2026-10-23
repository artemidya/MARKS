Private Sub AddCollectionToCollection(Col1 As Collection, col2 As Collection)
    Dim q As Long
    For q = 1 To col2.count
        Col1.Add col2(q)
    Next
End Sub
