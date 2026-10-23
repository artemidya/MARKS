Public Property Get SpeedsCount() As Long
    If Not Not lSpeeds Then
        SpeedsCount = UBound(lSpeeds) + 1
    End If
End Property
