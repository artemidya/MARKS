Public Function LongArrayToString(arr() As Long) As String
    If Not Not arr Then
        Dim q As Long
        For q = 0 To UBound(arr) - 1
            LongArrayToString = LongArrayToString & arr(q) & ";"
        Next
        LongArrayToString = LongArrayToString & arr(q)
    End If
End Function
