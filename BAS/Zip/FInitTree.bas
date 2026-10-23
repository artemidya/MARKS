Private Sub FInitTree()

  Dim i As Integer

    For i = n + 1 To (n + 256)
        rson(i) = NIL
    Next i
    For i = 0 To n - 1
        dad(i) = NIL
    Next i

End Sub
