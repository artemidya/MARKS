Public Property Get AccelerationsAsString() As String
    If Not Not lAccelerations Then
        Dim q As Long
        For q = 0 To UBound(lAccelerations) - 1
            AccelerationsAsString = AccelerationsAsString & lAccelerations(q) / 10 & ";"
        Next
        AccelerationsAsString = AccelerationsAsString & lAccelerations(q) / 10
        AccelerationsAsString = Replace(AccelerationsAsString, ",", ".")
    End If
End Property
