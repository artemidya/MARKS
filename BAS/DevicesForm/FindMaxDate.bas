Private Function FindMaxDate(pi() As PlotterInfo) As Long
    If Not Not pi Then
        Dim q As Long
        For q = 0 To UBound(pi)
            If pi(q).Date > FindMaxDate Then FindMaxDate = pi(q).Date
        Next
    End If
End Function
