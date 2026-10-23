Private Sub Sort(arr() As deviceInfo)
    If Not Not arr Then
        Dim i As Long
        Dim fSwitched As Boolean
        Dim arrTemp As deviceInfo
        Do
            fSwitched = False
            For i = 0 To UBound(arr) - 1
                If addZerroForDigits(arr(i).name) > addZerroForDigits(arr(i + 1).name) Then
                    fSwitched = True
                    arrTemp = arr(i)
                    arr(i) = arr(i + 1)
                    arr(i + 1) = arrTemp
                End If
            Next
        Loop While fSwitched
    End If
End Sub
