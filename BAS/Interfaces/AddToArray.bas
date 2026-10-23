Private Sub AddToArray(ByRef arr() As deviceInfo, newItem As deviceInfo)
    Dim NewIndex As Long
    If Not Not arr() Then
        NewIndex = UBound(arr) + 1
        ReDim Preserve arr(NewIndex)
        arr(NewIndex) = newItem
    Else
        ReDim arr(0)
        arr(0) = newItem
    End If
End Sub
