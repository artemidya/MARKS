Private Function StatusFromListByName(ByRef name As String, ByRef ConnectionStatuses() As ConnectionStatus) As Boolean
    Dim q As Long
    If Not Not ConnectionStatuses() Then
        For q = 0 To UBound(ConnectionStatuses())
            If name = ConnectionStatuses(q).name Then
                StatusFromListByName = ConnectionStatuses(q).Connected
                Exit Function
            End If
        Next
    End If
End Function
