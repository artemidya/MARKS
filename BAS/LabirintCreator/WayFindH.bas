Private Sub WayFindH(ByVal q As Long, ByRef Curves() As curve)
    Dim qq As Long, q2 As Long, s As Long, sLast As Long
    Dim ymax As Single, minR As Single, R As Single
    Dim ymin As Single
    Dim x As Single, y As Single
    Dim xD As Single, yD As Single
    Dim xUp As Single, yUp As Single
    Dim x2 As Single, y2 As Single
    Dim sx As Single, sy As Single
    Dim theEnd As Boolean ' Конец работы процедуры
    Dim toDown As Boolean ' Идти вниз а не вправо
    Dim toBegin As Boolean

    x = Curves(q).PositionX - 10
    y = Curves(q).PositionY - Curves(q).Height
    xD = x
    yD = y
    xUp = x
    yUp = y
    ymax = -99999999999999#
    ymin = 99999999999999#
    sLast = 0
    toBegin = True
10
    theEnd = True
    minR = 99999999999999#
    s = 0
    'Первый Проход
    For q2 = 1 To Curves(q).CountSubPath
        With Curves(q).SubPath(q2)
            If .IsSpeeding = False And .IsInner = False Then
                theEnd = False
                y2 = .LeftPosY
                If ymax > y2 Or sLast = 0 Then
                    x2 = .LeftPosX
                    If x2 >= x Then
                        R = Sqr((xUp - x2) ^ 2 + (yUp - y2) ^ 2)
                        If R < minR Then
                            minR = R
                            s = q2
                        End If
                    End If
                End If
            End If
        End With
    Next
    'Второй проход=============
    toDown = False
    If sLast <> 0 Then
        With Curves(q)
            sx = .SubPath(sLast).RightPosX
            sy = .SubPath(sLast).RightPosY
            minR = 99999999999999#
            For q2 = 1 To .CountSubPath
                With .SubPath(q2)
                    If .IsSpeeding = False And .IsInner = False Then
                        y2 = .TopPosY
                        x2 = .TopPosX
                        If y2 < sy And .PositionY > ymin Then
                            If x2 < sx Then
                                R = Sqr((xUp - x2) ^ 2 + (yUp - y2) ^ 2)
                                If R < minR Then
                                    minR = R
                                    s = q2
                                    toDown = True
                             End If
                            End If
                        End If
                    End If
                End With
            Next
        End With
    End If
    '===========================
20
    If s <> 0 Then
        If sLast <> 0 Then
            If toDown = False Then
                With Curves(q).SubPath(s)
                    x2 = .LeftPosX
                    y2 = .LeftPosY
                End With
                If GlobalCrossTest(x, y, x2, y2, Curves()) = False Then
                    AddLine x, y, x2, y2, Curves()
                    PointsH(CountPointsH).RightQ = q
                    PointsH(CountPointsH).RightQ2 = s
                Else
                    toBegin = True
                End If
            Else
                With Curves(q).SubPath(s)
                    x2 = .TopPosX
                    y2 = .TopPosY
                End With
                If GlobalCrossTest(xD, yD, x2, y2, Curves()) = False Then
                    AddLine xD, yD, x2, y2, Curves()
                    PointsH(CountPointsH).RightQ = q
                    PointsH(CountPointsH).RightQ2 = s
                Else
                    toBegin = True
                End If
            End If
        End If
        If toBegin = True Then
            CountPointsH = CountPointsH + 1
            ReDim Preserve PointsH(CountPointsH)
            With PointsH(CountPointsH)
                .LeftQ = q
                .LeftQ2 = s
                .RightQ = q
                .RightQ2 = s
            End With
            toBegin = False
        End If
        With Curves(q).SubPath(s)
             .IsSpeeding = True
            If .PositionY > ymax Then ymax = .PositionY: yUp = ymax
            If .PositionY - .Height < ymin Then ymin = .PositionY - .Height
            
            x = .RightPosX
            y = .RightPosY
            xD = .DownPosX
            yD = .DownPosY
            xUp = .PositionX + .With
        End With
    Else
        toBegin = True
        x = Curves(q).PositionX - 10
        y = Curves(q).PositionY - Curves(q).Height
        xD = x
        yD = y
        xUp = x
        yUp = y
        ymax = -99999999999999#
        ymin = 99999999999999#
    End If
   
    sLast = s
    If theEnd = True Then Exit Sub
    GoTo 10
End Sub
