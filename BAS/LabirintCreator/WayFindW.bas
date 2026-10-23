Private Sub WayFindW(ByVal q As Long, ByRef Curves() As curve)
    'Было   : Стало
    'Вправо - Вверх
    'Влево  - Вниз
    'Вниз   - Вправо
    'Вверх  - Влево
    Dim qq As Long, q2 As Long, s As Long, sLast As Long
    Dim xmax As Single, minR As Single, R As Single
    Dim xmin As Single
    Dim x As Single, y As Single
    Dim xR As Single, yR As Single
    Dim xL As Single, yL As Single
    Dim x2 As Single, y2 As Single
    Dim sx As Single, sy As Single

    'Обр
    Dim theEnd As Boolean ' Конец работы процедуры
    Dim toRight As Boolean ' Идти вниз а не вправо
    Dim toBegin As Boolean
    'Обр

    x = Curves(q).PositionX + Curves(q).With
    y = Curves(q).PositionY - Curves(q).Height - 10
    xR = x
    yR = y
    xL = x
    yL = y
    xmax = 99999999999999#
    xmin = -99999999999999#
    sLast = 0
    toBegin = True
    'Обр

10
    theEnd = True
    minR = 99999999999999#
    s = 0
    'Обр
    'Первый Проход
    For q2 = 1 To Curves(q).CountSubPath
        With Curves(q).SubPath(q2)
            If .IsSpeeding = False And .IsInner = False Then
            theEnd = False
                x2 = .DownPosX
                If xmax < x2 Or sLast = 0 Then
                    y2 = .DownPosY
                    If y2 >= y Then
                        R = Sqr((xL - x2) ^ 2 + (yL - y2) ^ 2)
                        If R < minR Then
                            minR = R
                            s = q2
                        End If
                    End If
                End If
            End If
        End With
    Next
    'Обр
    'Второй проход=============
    toRight = False
    'Обр
    If sLast <> 0 Then
        sx = Curves(q).SubPath(sLast).TopPosX
        sy = Curves(q).SubPath(sLast).TopPosY
        minR = 99999999999999#
    'Обр
        For q2 = 1 To Curves(q).CountSubPath
            With Curves(q).SubPath(q2)
                If .IsSpeeding = False And .IsInner = False Then
                    y2 = .LeftPosY
                    x2 = .LeftPosX
                    If x2 > sx And xmin > .PositionX Then
                        If y2 < sy Then
                            R = Sqr((xL - x2) ^ 2 + (yL - y2) ^ 2)
                            If R < minR Then
                                minR = R
                                s = q2
                                toRight = True
                            End If
                        End If
                    End If
                End If
            End With
        Next
    End If
    '===========================
    'Обр
20
    If s <> 0 Then
         If sLast <> 0 Then
             If toRight = False Then
                 With Curves(q).SubPath(s)
                    x2 = .DownPosX
                    y2 = .DownPosY
                 End With
                 If GlobalCrossTest(x, y, x2, y2, Curves()) = False Then
                     AddLine x, y, x2, y2, Curves()
                     PointsW(CountPointsW).TopQ = q
                     PointsW(CountPointsW).TopQ2 = s
                 Else
                     toBegin = True
                 End If
                 'Обр
             Else
                 With Curves(q).SubPath(s)
                    x2 = .LeftPosX
                    y2 = .LeftPosY
                 End With
                 If GlobalCrossTest(xR, yR, x2, y2, Curves()) = False Then
                     AddLine xR, yR, x2, y2, Curves()
                     PointsW(CountPointsW).TopQ = q
                     PointsW(CountPointsW).TopQ2 = s
                 Else
                     toBegin = True
                 End If
                 'Обр
             End If
         End If
         'Обр
         If toBegin = True Then
             CountPointsW = CountPointsW + 1
             ReDim Preserve PointsW(CountPointsW)
             PointsW(CountPointsW).DownQ = q
             PointsW(CountPointsW).DownQ2 = s
             PointsW(CountPointsW).TopQ = q
             PointsW(CountPointsW).TopQ2 = s
             toBegin = False
         End If
         
         With Curves(q).SubPath(s)
             'Обр
              .IsSpeeding = True
             If .PositionX < xmax Then xmax = .PositionX: xL = xmax
             If .PositionX + .With > xmin Then xmin = .PositionX + .With
             'Обр
            
             x = .TopPosX
             y = .TopPosY
             xR = .RightPosX
             yR = .RightPosY
             'Обр
             yL = .PositionY
             'Обр
         End With
    Else
        toBegin = True
        x = Curves(q).PositionX + Curves(q).With
        y = Curves(q).PositionY - Curves(q).Height - 10
        'Обр
        xR = x
        yR = y
        xL = x
        yL = y
        xmax = 99999999999999#
        xmin = -99999999999999#
    End If
   'Обр
    sLast = s
    If theEnd = True Then Exit Sub
    GoTo 10
End Sub
