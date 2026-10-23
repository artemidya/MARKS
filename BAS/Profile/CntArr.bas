Private Function CntArr(ByRef arr() As Long) As Long
    If Not Not arr Then
        CntArr = UBound(arr) + 1
    End If
End Function
