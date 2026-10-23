Private Sub IsInner(ByRef Curves() As curve)
    Dim q As Long, q2 As Long, qq As Long
    Dim x As Single, y As Single
    Dim x2 As Single, y2 As Single
    Dim rX As Single, dy As Single
    Dim rx2 As Single, dY2 As Single

    For q = 1 To CountCurves
        With Curves(q)
            For q2 = 1 To .CountSubPath
                With .SubPath(q2)
                    x2 = .PositionX
                    y2 = .PositionY
                    rx2 = x2 + .With
                    dY2 = y2 - .Height
                End With
                For qq = 1 To .CountSubPath
                    If qq <> q2 Then
                        With .SubPath(qq)
                            x = .PositionX
                            y = .PositionY
                            rX = x + .With
                            dy = y - .Height
                        End With
                        If x2 > x And rx2 < rX And y2 < y And dY2 > dy Then
                            .SubPath(q2).IsInner = True
                            Exit For
                        End If
                    End If
                Next
            Next
        End With
    Next
End Sub
