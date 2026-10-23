Public Property Let Item(Key As String, value As String)
    Dim q As Long
    For q = 0 To count - 1
        If Key = Keys(q) Then
            Values(q) = value
            Exit Property
        End If
    Next
    MsgBox "err"
End Property
