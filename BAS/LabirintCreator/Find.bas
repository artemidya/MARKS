Private Sub Find(ByRef Curves() As curve)
    CountPointsW = 0
    CountPointsH = 0
    Dim q As Long
    For q = 1 To CountCurves
        If Curves(q).With * 1.1 > Curves(q).Height Or Curves(q).CountSubPath < 3 Then
            WayFindH q, Curves()
        Else
            WayFindW q, Curves()
        End If
    Next
End Sub
