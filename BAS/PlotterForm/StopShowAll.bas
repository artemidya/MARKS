Private Sub StopShowAll()
    Dim q As Long
    For q = 0 To 7
        If Not PltDocument1(q) Is Nothing Then
            PltDocument1(q).StopShowWay
        End If
    Next
End Sub
