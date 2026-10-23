Private Sub Tops(ByRef Curves() As curve) 'Определяет топы для каждого подпути
    Dim q As Long, q2 As Long, qP As Long
    Dim x1 As Single, y1 As Single
    Dim x2 As Single, y2 As Single
    Dim xp As Single, yp As Single

    For q = 1 To CountCurves
        With Curves(q)
            For q2 = 1 To .CountSubPath
                With .SubPath(q2)
                    If .IsSpeeding = False And .IsInner = False Then
                        'Определяем левый ТОП =========================================
                        x1 = .PositionX - 1
                        x2 = .PositionX + .With + 1
                        y1 = .PositionY - .Height / 2
                        TopsCrossTest xp, yp, qP, x1, y1, x2, y1, Curves(q).SubPath(q2)
                        y1 = .PosY(qP)
                        y2 = .PosY(qP + 1)
                        If Abs(yp - y1) <= Abs(yp - y2) Then
                            .LeftPosX = .PosX(qP)
                            .LeftPosY = .PosY(qP)
                        Else
                            .LeftPosX = .PosX(qP + 1)
                            .LeftPosY = .PosY(qP + 1)
                        End If
                        'Определяем правый топ =========================================
                        y1 = .PositionY - .Height / 2
                        TopsCrossTest xp, yp, qP, x2, y1, x1, y1, Curves(q).SubPath(q2)
                        y1 = .PosY(qP)
                        y2 = .PosY(qP + 1)
                        If Abs(yp - y1) <= Abs(yp - y2) Then
                            .RightPosX = .PosX(qP)
                            .RightPosY = .PosY(qP)
                        Else
                            .RightPosX = .PosX(qP + 1)
                            .RightPosY = .PosY(qP + 1)
                        End If
        
                        'Определяем верхний ТОП =========================================
                        x1 = .PositionX + .With / 2
                        y1 = .PositionY + 1
                        y2 = .PositionY - .Height - 1
                        TopsCrossTest xp, yp, qP, x1, y1, x1, y2, Curves(q).SubPath(q2)
                        x1 = .PosX(qP)
                        x2 = .PosX(qP + 1)
                        '============
                        If Abs(xp - x1) <= Abs(xp - x2) Then
                            .TopPosX = .PosX(qP)
                            .TopPosY = .PosY(qP)
                        Else
                            .TopPosX = .PosX(qP + 1)
                            .TopPosY = .PosY(qP + 1)
                        End If
                        'Определяем нижний ТОП =========================================
                        x1 = .PositionX + .With / 2
                        TopsCrossTest xp, yp, qP, x1, y2, x1, y1, Curves(q).SubPath(q2)
                        x1 = .PosX(qP)
                        x2 = .PosX(qP + 1)
                        If Abs(xp - x1) <= Abs(xp - x2) Then
                            .DownPosX = .PosX(qP)
                            .DownPosY = .PosY(qP)
                        Else
                            .DownPosX = .PosX(qP + 1)
                            .DownPosY = .PosY(qP + 1)
                        End If
                    End If
                End With
            Next
        End With
    Next
End Sub
