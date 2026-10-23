Private Function CloneCollection(col As Collection) As Collection
    Set CloneCollection = New Collection
    Dim path As waPath
    Dim newPath As waPath
    For Each path In col
        Set newPath = path.Clone
        CloneCollection.Add newPath
    Next
End Function
