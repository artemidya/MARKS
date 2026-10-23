Private Function FindLong(ByVal value As Long, ByRef longBuffer() As Long) As Boolean
    Dim q As Long
    For q = 3 To UBound(longBuffer)
        If value = longBuffer(q) Then
            FindLong = True
            Exit Function
        End If
    Next
End Function
