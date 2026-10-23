Public Property Get Item(Key As String) As String
    Dim q As Long
    For q = 0 To count - 1
        If Key = Keys(q) Then
            Item = Values(q)
            Exit Function
        End If
    Next
End Property
