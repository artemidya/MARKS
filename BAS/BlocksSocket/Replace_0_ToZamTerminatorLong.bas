Private Sub Replace_0_ToZamTerminatorLong(ByVal value As Long, ByRef longBuffer() As Long)
    Dim q As Long
    For q = 3 To UBound(longBuffer)
        If 0 = longBuffer(q) Then longBuffer(q) = value
    Next
End Sub
