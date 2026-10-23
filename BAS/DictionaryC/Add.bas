Public Sub Add(Key As String, value As String)
    ReDim Preserve Keys(count)
    ReDim Preserve Values(count)
    Keys(count) = Key
    Values(count) = value
    count = count + 1
End Sub
