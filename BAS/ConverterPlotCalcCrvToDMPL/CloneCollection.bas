Private Function CloneCollection(col As Collection) As Collection
    Set CloneCollection = New Collection
    Dim path As waPath
    Dim newPath As waPath
    Dim q As Long
    For q = 1 To col.count
        Set path = col(q)
        Set newPath = path.Clone
        CloneCollection.Add newPath
    Next
End Function
