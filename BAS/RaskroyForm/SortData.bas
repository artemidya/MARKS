Private Sub SortData(ByRef x() As Double)
    If Not Not x() Then
        Dim i As Long
        Dim fSwitched As Boolean
        Dim xTemp As Double
        Do
            fSwitched = False
            For i = 1 To UBound(x) - 1
                If x(i) < x(i + 1) Then 'сортируем по возрастанию
                    fSwitched = True
                    xTemp = x(i)
                    x(i) = x(i + 1)
                    x(i + 1) = xTemp
                End If
            Next
        Loop While fSwitched
    End If
End Sub
