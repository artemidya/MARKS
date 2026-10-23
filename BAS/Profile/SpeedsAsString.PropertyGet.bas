Public Property Get SpeedsAsString() As String
    If Not Not lSpeeds Then
        Dim q As Long
        For q = 0 To UBound(lSpeeds) - 1
            SpeedsAsString = SpeedsAsString & lSpeeds(q) & ";"
        Next
        SpeedsAsString = SpeedsAsString & lSpeeds(q)
    End If
End Property
