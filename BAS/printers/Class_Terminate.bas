Private Sub Class_Terminate()
    Dim q As Long
    For q = 0 To lCount - 1
        If printers(q).Initialized Then DeInitialize (q)
    Next
End Sub
