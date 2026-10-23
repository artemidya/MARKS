Public Sub ErrShow()
    On Error GoTo 10
    Dim q As Long
    For q = 1 To UBound(ErrOnScreenCurves)
        ErrOnScreenCurves(q).Show
        ErrOnScreenText(q).Show
    Next
    Exit Sub
10:
End Sub
