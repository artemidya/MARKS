Private Function GetIndex(OutPutName As String) As Long
    Dim q As Long
    For q = 0 To UBound(Interfaces)
        If Interfaces(q).name = OutPutName Then
            GetIndex = q
            Exit Function
        End If
    Next
End Function
