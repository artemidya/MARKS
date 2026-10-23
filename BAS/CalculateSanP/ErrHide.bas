Public Sub ErrHide()
    On Error GoTo 10
    Dim q As Long
    For q = 1 To UBound(ErrOnScreenCurves)
        ErrOnScreenCurves(q).Hide
        ErrOnScreenText(q).Hide
    Next
    Exit Sub
10:
End Sub
