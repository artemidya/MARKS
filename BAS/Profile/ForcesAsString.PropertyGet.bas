Public Property Get ForcesAsString() As String
    If Not Not lForces Then
        Dim q As Long
        For q = 0 To UBound(lForces) - 1
            ForcesAsString = ForcesAsString & lForces(q) & ";"
        Next
        ForcesAsString = ForcesAsString & lForces(q)
    End If
End Property
