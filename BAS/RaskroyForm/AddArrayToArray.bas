Private Sub AddArrayToArray(ByRef arr() As MatrixForCheck, arr2() As MatrixForCheck)
    If Not Not arr() Then
        If Not Not arr2() Then
            Dim arr2count As Long
            Dim q As Long, c As Long
            arr2count = UBound(arr2()) + 1
            q = UBound(arr)
            ReDim Preserve arr(q + arr2count)
            For q = q + 1 To q + arr2count
                arr(q) = arr2(c)
                c = c + 1
            Next
        End If
    Else
        If Not Not arr2() Then
            arr() = arr2()
        End If
    End If
End Sub
