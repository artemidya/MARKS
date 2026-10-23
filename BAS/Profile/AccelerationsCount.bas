Public Property Get AccelerationsCount() As Long
    If Not Not lAccelerations Then
        AccelerationsCount = UBound(lAccelerations) + 1
        
    End If
End Property
