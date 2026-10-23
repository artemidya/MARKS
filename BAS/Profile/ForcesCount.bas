Public Property Get ForcesCount() As Long
    If Not Not lForces Then
        ForcesCount = UBound(lForces) + 1
    End If
End Property
