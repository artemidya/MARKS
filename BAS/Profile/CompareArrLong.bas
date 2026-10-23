Private Function CompareArrLong(ByRef Arr1() As Long, ByRef arr2() As Long) As Boolean
    If CntArr(Arr1) = CntArr(arr2) Then
        CompareArrLong = True
        If CntArr(Arr1) > 0 Then
            Dim q As Long
            For q = 0 To UBound(Arr1)
                If Arr1(q) <> arr2(q) Then
                    CompareArrLong = False
                    Exit For
                End If
            Next
        End If
    End If
End Function
