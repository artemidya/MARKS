' =============================================================================
' Group: 03_Geometry_and_Curves
' Source module: LabirintCreator.bas
' Procedures: 18
' =============================================================================

' -----------------------------------------------------------------------------
' CreateLabirintShape (Sub), source lines 58-90
' -----------------------------------------------------------------------------
Public Sub CreateLabirintShape()
    UserData.Load
    On Error GoTo 10
    
    If Application.ActiveSelection.shapes.count = 0 Then Exit Sub

    Dim lastUnit As cdrUnit
    lastUnit = Application.ActiveDocument.Unit
    Application.ActiveDocument.Unit = cdrInch
    
    Dim Curves() As curve
    
    Application.Optimization = True
    Application.ActiveSelection.ConvertToCurves
    Application.ActiveSelection.UngroupAllEx
    GetCurves Curves()          'Преобразовуем кривые для обработки
    ActiveDocument.Undo 2
    IsInner Curves()            'Определяет является ли подпуть внутренним
    Tops Curves()               'Определяет точки-соединения
    Find Curves()               'Локальное соединение
    GlobConnectH Curves()       'Глобальное горизонтальное соединение
    GlobConnectW Curves()       'Глобальное вертикальное соединение
    OverENDconnect_H Curves()   'Последнее горизонтальное соединение
    OverENDconnect_W Curves()   'Последнее вертикальное соединение
    DrawCurves Curves()         'нарисовать кривые
    
10
    Application.ActiveDocument.Unit = lastUnit
    Application.Optimization = False
    Application.Refresh
    

End Sub

' -----------------------------------------------------------------------------
' DrawCurves (Sub), source lines 91-114
' -----------------------------------------------------------------------------
Private Sub DrawCurves(ByRef Curves() As curve)
    Dim q As Long
    Dim crv As New VGCore.curve
    Dim shape As VGCore.shape
    Dim sr As ShapeRange
    Set sr = ActiveSelectionRange
    
    For q = CountCurves + 1 To CountLineQ
        With Curves(q).SubPath(1)
            crv.CreateSubPath(.PosX(1), .PosY(1)).AppendLineSegment .PosX(2), .PosY(2)
        End With
    Next
    'Чертит прямоугольник
    crv.AppendSubpathRectangle XSel, YSel, XSel + WithSel, YSel - HeightSel
    
    Application.ActiveDocument.BeginCommandGroup localization.ActiveLng.Item("LabirintCreator.BeginCommandGroup")
    Set shape = ActiveLayer.CreateCurve(crv)
    shape.Outline.SetProperties Color:=CreateCMYKColor(0, 60, 100, 0)
    
    sr.AddToSelection
    ActiveSelection.Group
    
    Application.ActiveDocument.EndCommandGroup
End Sub

' -----------------------------------------------------------------------------
' GetCurves (Sub), source lines 116-259
' -----------------------------------------------------------------------------
Private Sub GetCurves(ByRef Curves() As curve)
    Dim q As Long, q2 As Long, q3 As Long
    Dim maxX As Single, minX As Single
    Dim maxY As Single, minY As Single

    Dim p0x As Single, p0y As Single
    Dim p1x As Single, p1y As Single
    Dim p2x As Single, p2y As Single
    Dim p3x As Single, p3y As Single
    
    Dim dx As Single, dy As Single
    Dim leng As Single
    Dim stp As Single
    Dim a As Long, k As Long
    Dim t As Single
    Dim s As Long
    
    XSel = Application.ActiveSelection.PositionX - 0.12
    YSel = Application.ActiveSelection.PositionY + 0.12
    WithSel = Application.ActiveSelection.SizeWidth + 0.24
    HeightSel = Application.ActiveSelection.SizeHeight + 0.24
    
    CountCurves = Application.ActiveSelection.shapes.count
    CountLineQ = Application.ActiveSelection.shapes.count

    ReDim Curves(Application.ActiveSelection.shapes.count)
    
    For q = 1 To Application.ActiveSelection.shapes.count
        With Application.ActiveSelection.shapes(q)
            Curves(q).PositionX = .PositionX
            Curves(q).PositionY = .PositionY
            Curves(q).With = .SizeWidth
            Curves(q).Height = .SizeHeight
            If .Type = cdrCurveShape Then
                Curves(q).CountSubPath = .curve.Subpaths.count
                ReDim Curves(q).SubPath(.curve.Subpaths.count)
                
                For q2 = 1 To .curve.Subpaths.count
                    With .curve.Subpaths(q2)
                        Curves(q).SubPath(q2).PositionX = .PositionX
                        Curves(q).SubPath(q2).PositionY = .PositionY
                        Curves(q).SubPath(q2).CountNodes = .Nodes.count
                        a = 0
                        maxX = -100000000000#: minX = 100000000000#
                        maxY = -100000000000#: minY = 100000000000#
                        For q3 = 1 To .Nodes.count
                            p0x = .Nodes(q3).PositionX
                            p0y = .Nodes(q3).PositionY
                           
                            If q3 < .Nodes.count Then
                                p3x = .Nodes(q3 + 1).PositionX
                                p3y = .Nodes(q3 + 1).PositionY
                                
                                If .Nodes(q3 + 1).segment.Type = cdrLineSegment Then
                                    p1x = p0x
                                    p1y = p0y
                                    p2x = p3x
                                    p2y = p3y
                                Else
                                    p1x = .Nodes(q3 + 1).segment.StartingControlPointX
                                    p1y = .Nodes(q3 + 1).segment.StartingControlPointY
                                    p2x = .Nodes(q3 + 1).segment.EndingControlPointX
                                    p2y = .Nodes(q3 + 1).segment.EndingControlPointY
                                End If
                                
                            Else
                                p3x = .Nodes(1).PositionX
                                p3y = .Nodes(1).PositionY
                                
                                If .Nodes(1).segment.Type = cdrLineSegment Then
                                    p1x = p0x
                                    p1y = p0y
                                    p2x = p3x
                                    p2y = p3y
                                Else
                                    p1x = .Nodes(1).segment.StartingControlPointX
                                    p1y = .Nodes(1).segment.StartingControlPointY
                                    p2x = .Nodes(1).segment.EndingControlPointX
                                    p2y = .Nodes(1).segment.EndingControlPointY
                                End If
                                
                            End If
                            
                           
                            If q3 = .Nodes.count Then
                                dx = .Nodes(q3).PositionX - .Nodes(1).PositionX
                                dy = .Nodes(q3).PositionY - .Nodes(1).PositionY
                            Else
                                dx = .Nodes(q3).PositionX - .Nodes(q3 + 1).PositionX
                                dy = .Nodes(q3).PositionY - .Nodes(q3 + 1).PositionY
                            End If
                            leng = dx * dx + dy * dy
                            If .Nodes.count = 1 Then leng = 100
                            If leng = 0 Then GoTo 100
                            k = 10
                            'If k < 2 Then k = 2
                            'If k / 2 <> k \ 2 Then k = k - 1
                            stp = 0.1
                            a = a + 1
                            ReDim Preserve Curves(q).SubPath(q2).PosX(a + 10)
                            ReDim Preserve Curves(q).SubPath(q2).PosY(a + 10)
                            
                            Curves(q).SubPath(q2).PosX(a) = .Nodes(q3).PositionX
                            Curves(q).SubPath(q2).PosY(a) = .Nodes(q3).PositionY
                            With Curves(q).SubPath(q2)
                                If .PosX(a) > maxX Then maxX = .PosX(a)
                                If .PosX(a) < minX Then minX = .PosX(a)
                                If .PosY(a) > maxY Then maxY = .PosY(a)
                                If .PosY(a) < minY Then minY = .PosY(a)
                            
                                t = 0
                                For s = 1 To k - 1
                                    t = t + stp
                                    a = a + 1
                                    .PosX(a) = (1 - t) ^ 3 * p0x + 3 * (1 - t) ^ 2 * t * p1x + 3 * (1 - t) * t ^ 2 * p2x + t ^ 3 * p3x
                                    .PosY(a) = (1 - t) ^ 3 * p0y + 3 * (1 - t) ^ 2 * t * p1y + 3 * (1 - t) * t ^ 2 * p2y + t ^ 3 * p3y
                                    If .PosX(a) > maxX Then maxX = .PosX(a)
                                    If .PosX(a) < minX Then minX = .PosX(a)
                                    If .PosY(a) > maxY Then maxY = .PosY(a)
                                    If .PosY(a) < minY Then minY = .PosY(a)
                                Next
100
                            End With
                        Next
                        a = a + 1
                        ReDim Preserve Curves(q).SubPath(q2).PosX(a)
                        ReDim Preserve Curves(q).SubPath(q2).PosY(a)
                        Curves(q).SubPath(q2).PosX(a) = .Nodes(1).PositionX
                        Curves(q).SubPath(q2).PosY(a) = .Nodes(1).PositionY
                    End With
                    With Curves(q).SubPath(q2)
                        If .PosX(a) > maxX Then maxX = .PosX(a)
                        If .PosX(a) < minX Then minX = .PosX(a)
                        If .PosY(a) > maxY Then maxY = .PosY(a)
                        If .PosY(a) < minY Then minY = .PosY(a)
                        .With = maxX - minX
                        .Height = maxY - minY
                        .CountNodes = a
                    End With
                Next
            End If
        End With
    Next
End Sub

' -----------------------------------------------------------------------------
' IsInner (Sub), source lines 260-293
' -----------------------------------------------------------------------------
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

' -----------------------------------------------------------------------------
' Tops (Sub), source lines 294-364
' -----------------------------------------------------------------------------
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

' -----------------------------------------------------------------------------
' TopsCrossTest (Sub), source lines 365-411
' -----------------------------------------------------------------------------
Private Sub TopsCrossTest(ByRef xp As Single, ByRef yp As Single, ByRef qP As Long, ByVal x1 As Single, ByVal y1 As Single, ByVal x2 As Single, ByVal y2 As Single, ByRef SubPath As SubPath)
    'Для нахождения топов!
    Dim ua As Single, ub As Single
    Dim znamen As Single, chis1 As Single, chis2 As Single
    Dim x3 As Single, x4 As Single
    Dim y3 As Single, y4 As Single
    Dim xPtemp As Single, yPtemp As Single
    Dim minR As Single, R As Single
    Dim q3 As Long
    minR = 99999999999999#

    With SubPath
        For q3 = 1 To .CountNodes - 1
            x3 = .PosX(q3)
            x4 = .PosX(q3 + 1)
            If x3 > x1 And x4 > x1 And x3 > x2 And x4 > x2 Then GoTo 10
            If x3 < x1 And x4 < x1 And x3 < x2 And x4 < x2 Then GoTo 10
            
            y3 = .PosY(q3)
            y4 = .PosY(q3 + 1)
            If y3 > y1 And y4 > y1 And y3 > y2 And y4 > y2 Then GoTo 10
            If y3 < y1 And y4 < y1 And y3 < y2 And y4 < y2 Then GoTo 10
            
            znamen = (y4 - y3) * (x2 - x1) - (x4 - x3) * (y2 - y1)
            chis1 = (x4 - x3) * (y1 - y3) - (y4 - y3) * (x1 - x3)
            chis2 = (x2 - x1) * (y1 - y3) - (y2 - y1) * (x1 - x3)
        
            If znamen = 0 Then GoTo 10
            ua = chis1 / znamen
            ub = chis2 / znamen
            xPtemp = x1 + ua * (x2 - x1)
            yPtemp = y1 + ua * (y2 - y1)
            If ua >= 0 And ua <= 1 And ub >= 0 And ub <= 1 Then
                xPtemp = x1 + ua * (x2 - x1)
                yPtemp = y1 + ua * (y2 - y1)
                R = Sqr((xPtemp - x1) ^ 2 + (yPtemp - y1) ^ 2)
                If R < minR Then
                    xp = xPtemp
                    yp = yPtemp
                    qP = q3
                    minR = R
                End If
            End If
10
        Next
    End With
End Sub

' -----------------------------------------------------------------------------
' Find (Sub), source lines 412-423
' -----------------------------------------------------------------------------
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

' -----------------------------------------------------------------------------
' WayFindH (Sub), source lines 424-564
' -----------------------------------------------------------------------------
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

' -----------------------------------------------------------------------------
' WayFindW (Sub), source lines 565-725
' -----------------------------------------------------------------------------
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

' -----------------------------------------------------------------------------
' AddLine (Sub), source lines 726-761
' -----------------------------------------------------------------------------
Private Sub AddLine(ByVal x As Single, ByVal y As Single, ByVal x2 As Single, ByVal y2 As Single, ByRef Curves() As curve)
    'Процедура добавляет линию!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
    CountLineQ = CountLineQ + 1
    ReDim Preserve Curves(CountLineQ)
    With Curves(CountLineQ)
        .CountSubPath = 1
        ReDim .SubPath(1)
        If x < x2 Then
            .PositionX = x
            .With = x2 - x
        Else
            .PositionX = x2
            .With = x - x2
        End If
        .SubPath(1).PositionX = .PositionX
        .SubPath(1).With = .With
        If y > y2 Then
            .PositionY = y
            .Height = y - y2
        Else
            .PositionY = y2
            .Height = y2 - y
        End If
        .SubPath(1).PositionY = .PositionY
        .SubPath(1).Height = .Height
        With .SubPath(1)
            .CountNodes = 2
            ReDim .PosX(2)
            ReDim .PosY(2)
            .PosX(1) = x
            .PosY(1) = y
            .PosX(2) = x2
            .PosY(2) = y2
        End With
    End With
End Sub

' -----------------------------------------------------------------------------
' GlobalCrossTest (Function), source lines 762-821
' -----------------------------------------------------------------------------
Private Function GlobalCrossTest(ByVal x1 As Single, ByVal y1 As Single, ByVal x2 As Single, ByVal y2 As Single, ByRef Curves() As curve) As Boolean
    Dim ua As Single, ub As Single
    Dim znamen As Single, chis1 As Single, chis2 As Single
    Dim x3 As Single, x4 As Single
    Dim y3 As Single, y4 As Single
    Dim q As Long, q2 As Long, q3 As Long
    
    For q = 1 To CountLineQ
        With Curves(q)
            x3 = .PositionX
            x4 = .PositionX + .With
            If x3 > x1 And x4 > x1 And x3 > x2 And x4 > x2 Then GoTo 30
            If x3 < x1 And x4 < x1 And x3 < x2 And x4 < x2 Then GoTo 30
            
            y3 = .PositionY - .Height
            y4 = .PositionY
            If y3 > y1 And y4 > y1 And y3 > y2 And y4 > y2 Then GoTo 30
            If y3 < y1 And y4 < y1 And y3 < y2 And y4 < y2 Then GoTo 30
        
            For q2 = 1 To .CountSubPath
                With .SubPath(q2)
                    x3 = .PositionX
                    x4 = .PositionX + .With
                    If x3 > x1 And x4 > x1 And x3 > x2 And x4 > x2 Then GoTo 20
                    If x3 < x1 And x4 < x1 And x3 < x2 And x4 < x2 Then GoTo 20
                    
                    y3 = .PositionY - .Height
                    y4 = .PositionY
                    If y3 > y1 And y4 > y1 And y3 > y2 And y4 > y2 Then GoTo 20
                    If y3 < y1 And y4 < y1 And y3 < y2 And y4 < y2 Then GoTo 20
                
                    For q3 = 1 To .CountNodes - 1
                        x3 = .PosX(q3)
                        x4 = .PosX(q3 + 1)
                        If x3 > x1 And x4 > x1 And x3 > x2 And x4 > x2 Then GoTo 10
                        If x3 < x1 And x4 < x1 And x3 < x2 And x4 < x2 Then GoTo 10
                    
                        y3 = .PosY(q3)
                        y4 = .PosY(q3 + 1)
                        If y3 > y1 And y4 > y1 And y3 > y2 And y4 > y2 Then GoTo 10
                        If y3 < y1 And y4 < y1 And y3 < y2 And y4 < y2 Then GoTo 10
                    
                        znamen = (y4 - y3) * (x2 - x1) - (x4 - x3) * (y2 - y1)
                        chis1 = (x4 - x3) * (y1 - y3) - (y4 - y3) * (x1 - x3)
                        chis2 = (x2 - x1) * (y1 - y3) - (y2 - y1) * (x1 - x3)
                        If znamen = 0 Then GoTo 10
                    
                        ua = chis1 / znamen
                        ub = chis2 / znamen
                        If ua > 0 And ua < 1 And ub > 0 And ub < 1 Then GlobalCrossTest = True: Exit Function
10
                    Next
20
                End With
            Next
30
        End With
    Next
    GlobalCrossTest = False
End Function

' -----------------------------------------------------------------------------
' GlobConnectH (Sub), source lines 822-916
' -----------------------------------------------------------------------------
Private Sub GlobConnectH(ByRef Curves() As curve)
    'Глобальное горизонтальное соединение
    Dim q As Long, i As Long
    Dim tQ As Long, tQ2 As Long
    Dim t2Q As Long, t2Q2 As Long
    Dim s As Long, s2 As Long
    Dim a As Long
    Dim x As Single, y As Single
    Dim x2 As Single, y2 As Single
    Dim R As Single, minR As Single
    Dim xTemp As Single, yTemp As Single
    
    For q = 1 To CountPointsH
        If PointsH(q).IsRightStend = False Then
            s = 0
            minR = 99999999999999#
            tQ = PointsH(q).RightQ
            tQ2 = PointsH(q).RightQ2
            x = Curves(tQ).SubPath(tQ2).RightPosX
            y = Curves(tQ).SubPath(tQ2).RightPosY
            
            For i = 1 To CountPointsH
                If i <> q And PointsH(i).IsLeftStend = False Then
                    t2Q = PointsH(i).LeftQ
                    t2Q2 = PointsH(i).LeftQ2
                    x2 = Curves(t2Q).SubPath(t2Q2).LeftPosX
                    If x2 > x Then
                        y2 = Curves(t2Q).SubPath(t2Q2).LeftPosY
                        R = Sqr((x - x2) ^ 2 + (y - y2) ^ 2)
                        If R < minR Then
                            minR = R
                            s = i
                        End If
                    End If
                End If
            Next
        End If
        If s <> 0 Then
            t2Q = PointsH(s).LeftQ
            t2Q2 = PointsH(s).LeftQ2
            x2 = Curves(t2Q).SubPath(t2Q2).LeftPosX
            y2 = Curves(t2Q).SubPath(t2Q2).LeftPosY
        
            If LocalCrossTest(x, y, x2, y2, Curves(tQ).SubPath(tQ2)) = True Then
                minR = 999999999999#
                s2 = 0
                xTemp = Curves(tQ).SubPath(tQ2).RightPosX
                yTemp = Curves(tQ).SubPath(tQ2).RightPosY
                For a = 1 To Curves(tQ).SubPath(tQ2).CountNodes
                    x = Curves(tQ).SubPath(tQ2).PosX(a)
                    y = Curves(tQ).SubPath(tQ2).PosY(a)
                    If LocalCrossTest(x, y, x2, y2, Curves(tQ).SubPath(tQ2)) = False Then
                        R = Sqr((x - xTemp) ^ 2 + (y - yTemp) ^ 2)
                        If R < minR Then
                            minR = R
                            s2 = a
                        End If
                    End If
                Next
                If s2 <> 0 Then
                    x = Curves(tQ).SubPath(tQ2).PosX(s2)
                    y = Curves(tQ).SubPath(tQ2).PosY(s2)
                End If
            End If
                
            If LocalCrossTest(x, y, x2, y2, Curves(t2Q).SubPath(t2Q2)) = True Then
                minR = 999999999999#
                s2 = 0
                xTemp = Curves(t2Q).SubPath(t2Q2).LeftPosX
                yTemp = Curves(t2Q).SubPath(t2Q2).LeftPosY
                For a = 1 To Curves(t2Q).SubPath(t2Q2).CountNodes
                    x2 = Curves(t2Q).SubPath(t2Q2).PosX(a)
                    y2 = Curves(t2Q).SubPath(t2Q2).PosY(a)
                    If LocalCrossTest(x, y, x2, y2, Curves(t2Q).SubPath(t2Q2)) = False And LocalCrossTest(x, y, x2, y2, Curves(tQ).SubPath(tQ2)) = False Then
                        R = Sqr((xTemp - x2) ^ 2 + (yTemp - y2) ^ 2)
                        If R < minR Then
                        minR = R
                        s2 = a
                        End If
                    End If
                Next
                If s2 <> 0 Then
                    x2 = Curves(t2Q).SubPath(t2Q2).PosX(s2)
                    y2 = Curves(t2Q).SubPath(t2Q2).PosY(s2)
                End If
            End If
            
            If GlobalCrossTest(x, y, x2, y2, Curves()) = False Then
                AddLine x, y, x2, y2, Curves()
                PointsH(s).IsLeftStend = True
                PointsH(q).IsRightStend = True
            End If
        End If
    Next
End Sub

' -----------------------------------------------------------------------------
' GlobConnectW (Sub), source lines 917-1019
' -----------------------------------------------------------------------------
Private Sub GlobConnectW(ByRef Curves() As curve)
    'Глобальное вертикальное соединение
    ' Было  : Стало
    'Вправо - Вверх
    'Влево  - Вниз
    'Вниз   - Вправо
    'Вверх  - Влево
    Dim q As Long, i As Long
    Dim tQ As Long, tQ2 As Long
    Dim t2Q As Long, t2Q2 As Long
    Dim s As Long, s2 As Long
    Dim a As Long
    
    Dim x As Single, y As Single
    Dim x2 As Single, y2 As Single
    Dim R As Single, minR As Single
    Dim xTemp As Single, yTemp As Single
    
    For q = 1 To CountPointsW
        If PointsW(q).IsTopStend = False Then
            s = 0
            minR = 99999999999999#
            tQ = PointsW(q).TopQ
            tQ2 = PointsW(q).TopQ2
            x = Curves(tQ).SubPath(tQ2).TopPosX
            y = Curves(tQ).SubPath(tQ2).TopPosY

            For i = 1 To CountPointsW
                If i <> q And PointsW(i).IsDownStend = False Then
                    t2Q = PointsW(i).DownQ
                    t2Q2 = PointsW(i).DownQ2
                    y2 = Curves(t2Q).SubPath(t2Q2).DownPosY
                    If y2 > y Then
                        x2 = Curves(t2Q).SubPath(t2Q2).DownPosX
                        R = Sqr((x - x2) ^ 2 + (y - y2) ^ 2)
                        If R < minR Then
                            minR = R
                            s = i
                        End If
                    End If
                End If
            Next
        End If
        
        If s <> 0 Then
            t2Q = PointsW(s).DownQ
            t2Q2 = PointsW(s).DownQ2
            x2 = Curves(t2Q).SubPath(t2Q2).DownPosX
            y2 = Curves(t2Q).SubPath(t2Q2).DownPosY
            
            If LocalCrossTest(x, y, x2, y2, Curves(tQ).SubPath(tQ2)) = True Then
                 minR = 999999999999#
                 s2 = 0
                 xTemp = Curves(tQ).SubPath(tQ2).TopPosX
                 yTemp = Curves(tQ).SubPath(tQ2).TopPosY
                For a = 1 To Curves(tQ).SubPath(tQ2).CountNodes
                    x = Curves(tQ).SubPath(tQ2).PosX(a)
                    y = Curves(tQ).SubPath(tQ2).PosY(a)
                    If LocalCrossTest(x, y, x2, y2, Curves(tQ).SubPath(tQ2)) = False Then
                        R = Sqr((x - xTemp) ^ 2 + (y - yTemp) ^ 2)
                        If R < minR Then
                        minR = R
                        s2 = a
                        End If
                    End If
                Next
                    If s2 <> 0 Then
                    x = Curves(tQ).SubPath(tQ2).PosX(s2)
                    y = Curves(tQ).SubPath(tQ2).PosY(s2)
                    End If
            End If

            If LocalCrossTest(x, y, x2, y2, Curves(t2Q).SubPath(t2Q2)) = True Then
                minR = 999999999999#
                s2 = 0
                xTemp = Curves(t2Q).SubPath(t2Q2).DownPosX
                yTemp = Curves(t2Q).SubPath(t2Q2).DownPosY
                For a = 1 To Curves(t2Q).SubPath(t2Q2).CountNodes
                    x2 = Curves(t2Q).SubPath(t2Q2).PosX(a)
                    y2 = Curves(t2Q).SubPath(t2Q2).PosY(a)
                    If LocalCrossTest(x, y, x2, y2, Curves(t2Q).SubPath(t2Q2)) = False And LocalCrossTest(x, y, x2, y2, Curves(tQ).SubPath(tQ2)) = False Then
                        R = Sqr((xTemp - x2) ^ 2 + (yTemp - y2) ^ 2)
                        If R < minR Then
                            minR = R
                            s2 = a
                        End If
                    End If
                Next
                If s2 <> 0 Then
                    x2 = Curves(t2Q).SubPath(t2Q2).PosX(s2)
                    y2 = Curves(t2Q).SubPath(t2Q2).PosY(s2)
                End If
            End If
               
            If GlobalCrossTest(x, y, x2, y2, Curves()) = False Then
                'Set s1 = ActiveLayer.CreateLineSegment(x, y, x2, y2)
                AddLine x, y, x2, y2, Curves()
                PointsW(s).IsDownStend = True
                PointsW(q).IsTopStend = True
            End If
        End If
    Next
End Sub

' -----------------------------------------------------------------------------
' LocalCrossTest (Function), source lines 1020-1054
' -----------------------------------------------------------------------------
Private Function LocalCrossTest(ByVal x1 As Single, ByVal y1 As Single, ByVal x2 As Single, ByVal y2 As Single, ByRef SubPath As SubPath) As Boolean
    Dim ua As Single, ub As Single
    Dim znamen As Single, chis1 As Single, chis2 As Single
    Dim x3 As Single, x4 As Single
    Dim y3 As Single, y4 As Single
    Dim q3 As Long
    With SubPath
        For q3 = 1 To .CountNodes - 1
            x3 = .PosX(q3)
            x4 = .PosX(q3 + 1)
            If x3 > x1 And x4 > x1 And x3 > x2 And x4 > x2 Then GoTo 10
            If x3 < x1 And x4 < x1 And x3 < x2 And x4 < x2 Then GoTo 10
            
            y3 = .PosY(q3)
            y4 = .PosY(q3 + 1)
            If y3 > y1 And y4 > y1 And y3 > y2 And y4 > y2 Then GoTo 10
            If y3 < y1 And y4 < y1 And y3 < y2 And y4 < y2 Then GoTo 10
            
            znamen = (y4 - y3) * (x2 - x1) - (x4 - x3) * (y2 - y1)
            chis1 = (x4 - x3) * (y1 - y3) - (y4 - y3) * (x1 - x3)
            chis2 = (x2 - x1) * (y1 - y3) - (y2 - y1) * (x1 - x3)
            
            If znamen = 0 Then
                If chis1 = 0 Or chis2 = 0 Then LocalCrossTest = True: GoTo 20
            Else
                ua = chis1 / znamen
                ub = chis2 / znamen
                If ua > 0 And ua < 1 And ub > 0 And ub < 1 Then LocalCrossTest = True: GoTo 20
            End If
10
        Next
        LocalCrossTest = False
20
    End With
End Function

' -----------------------------------------------------------------------------
' OverENDconnect_H (Sub), source lines 1055-1088
' -----------------------------------------------------------------------------
Private Sub OverENDconnect_H(ByRef Curves() As curve)
    'Последнее горизонтальное соединение
    Dim x1 As Single, y1 As Single
    Dim x2 As Single, y2 As Single
    Dim xp As Single, yp As Single
    Dim q As Long, q1 As Long, q2 As Long
    Dim SecError As Boolean
    
    For q = 1 To CountPointsH
        If PointsH(q).IsLeftStend = False Then
            q1 = PointsH(q).LeftQ
            q2 = PointsH(q).LeftQ2
            x1 = Curves(q1).SubPath(q2).LeftPosX
            y1 = Curves(q1).SubPath(q2).LeftPosY
            If OverCrossTest(xp, yp, x1, y1, XSel, y1, SecError, Curves()) = True Then
                If SecError = False Then AddLine x1, y1, xp, yp, Curves()               ' Добавляем линию
            Else
                AddLine x1, y1, XSel, y1, Curves() ' Добавляем линию до конца
            End If
        End If
        '=====================
        If PointsH(q).IsRightStend = False Then
            q1 = PointsH(q).RightQ
            q2 = PointsH(q).RightQ2
            x1 = Curves(q1).SubPath(q2).RightPosX
            y1 = Curves(q1).SubPath(q2).RightPosY
            If OverCrossTest(xp, yp, x1, y1, XSel + WithSel, y1, SecError, Curves()) = True Then
                If SecError = False Then AddLine x1, y1, xp, yp, Curves() ' Добавляем линию
            Else
                AddLine x1, y1, XSel + WithSel, y1, Curves() ' Добавляем линию до конца
            End If
        End If
    Next
End Sub

' -----------------------------------------------------------------------------
' OverENDconnect_W (Sub), source lines 1089-1122
' -----------------------------------------------------------------------------
Private Sub OverENDconnect_W(ByRef Curves() As curve)
    'Последнее вертикальное соединение
    Dim x1 As Single, y1 As Single
    Dim x2 As Single, y2 As Single
    Dim xp As Single, yp As Single
    Dim q As Long, q1 As Long, q2 As Long
    Dim SecError As Boolean
    
    For q = 1 To CountPointsW
        If PointsW(q).IsTopStend = False Then
            q1 = PointsW(q).TopQ
            q2 = PointsW(q).TopQ2
            x1 = Curves(q1).SubPath(q2).TopPosX
            y1 = Curves(q1).SubPath(q2).TopPosY
            If OverCrossTest(xp, yp, x1, y1, x1, YSel, SecError, Curves()) = True Then
                If SecError = False Then AddLine x1, y1, xp, yp, Curves() ' Добавляем линию
            Else
                AddLine x1, y1, x1, YSel, Curves() ' Добавляем линию до конца
            End If
        End If
        '=====================
        If PointsW(q).IsDownStend = False Then
            q1 = PointsW(q).DownQ
            q2 = PointsW(q).DownQ2
            x1 = Curves(q1).SubPath(q2).DownPosX
            y1 = Curves(q1).SubPath(q2).DownPosY
            If OverCrossTest(xp, yp, x1, y1, x1, YSel - HeightSel, SecError, Curves()) = True Then
                If SecError = False Then AddLine x1, y1, xp, yp, Curves() ' Добавляем линию
            Else
                AddLine x1, y1, x1, YSel - HeightSel, Curves() ' Добавляем линию до конца
            End If
        End If
    Next
End Sub

' -----------------------------------------------------------------------------
' OverCrossTest (Function), source lines 1123-1222
' -----------------------------------------------------------------------------
Private Function OverCrossTest(ByRef xp As Single, ByRef yp As Single, ByVal x1 As Single, ByVal y1 As Single, ByVal x2 As Single, ByVal y2 As Single, ByRef SecError As Boolean, ByRef Curves() As curve) As Boolean
    SecError = False
    'Процедура для нахождения пересечений при окончании работы.
    'xP возвращается X пересечения
    'yP возвращается Y пересечения
    xp = 0
    yp = 0
    OverCrossTest = False
    Dim ua As Single, ub As Single
    Dim znamen As Single, chis1 As Single, chis2 As Single
    Dim x3 As Single, x4 As Single
    Dim y3 As Single, y4 As Single
    Dim minR As Single, R As Single
    Dim r1 As Single, R2 As Single
    Dim xPtemp As Single, yPtemp As Single
    minR = 99999999999999#
    Dim q As Long, q2 As Long, q3 As Long
    Dim qSec As Long, q2Sec As Long
    
    For q = 1 To CountLineQ
        x3 = Curves(q).PositionX
        x4 = Curves(q).PositionX + Curves(q).With
        If x3 > x1 And x4 > x1 And x3 > x2 And x4 > x2 Then GoTo 30
        If x3 < x1 And x4 < x1 And x3 < x2 And x4 < x2 Then GoTo 30
        
        y3 = Curves(q).PositionY - Curves(q).Height
        y4 = Curves(q).PositionY
        If y3 > y1 And y4 > y1 And y3 > y2 And y4 > y2 Then GoTo 30
        If y3 < y1 And y4 < y1 And y3 < y2 And y4 < y2 Then GoTo 30
        
        For q2 = 1 To Curves(q).CountSubPath
            x3 = Curves(q).SubPath(q2).PositionX
            x4 = Curves(q).SubPath(q2).PositionX + Curves(q).SubPath(q2).With
            If x3 > x1 And x4 > x1 And x3 > x2 And x4 > x2 Then GoTo 20
            If x3 < x1 And x4 < x1 And x3 < x2 And x4 < x2 Then GoTo 20
            
            y3 = Curves(q).SubPath(q2).PositionY - Curves(q).SubPath(q2).Height
            y4 = Curves(q).SubPath(q2).PositionY
            If y3 > y1 And y4 > y1 And y3 > y2 And y4 > y2 Then GoTo 20
            If y3 < y1 And y4 < y1 And y3 < y2 And y4 < y2 Then GoTo 20
        
            For q3 = 1 To Curves(q).SubPath(q2).CountNodes - 1
                x3 = Curves(q).SubPath(q2).PosX(q3)
                x4 = Curves(q).SubPath(q2).PosX(q3 + 1)
                If x3 > x1 And x4 > x1 And x3 > x2 And x4 > x2 Then GoTo 10
                If x3 < x1 And x4 < x1 And x3 < x2 And x4 < x2 Then GoTo 10
                
                y3 = Curves(q).SubPath(q2).PosY(q3)
                y4 = Curves(q).SubPath(q2).PosY(q3 + 1)
                If y3 > y1 And y4 > y1 And y3 > y2 And y4 > y2 Then GoTo 10
                If y3 < y1 And y4 < y1 And y3 < y2 And y4 < y2 Then GoTo 10
                
                znamen = (y4 - y3) * (x2 - x1) - (x4 - x3) * (y2 - y1)
                chis1 = (x4 - x3) * (y1 - y3) - (y4 - y3) * (x1 - x3)
                chis2 = (x2 - x1) * (y1 - y3) - (y2 - y1) * (x1 - x3)
                If znamen = 0 Then GoTo 10
                
                ua = chis1 / znamen
                ub = chis2 / znamen
                If ua >= 0 And ua <= 1 And ub >= 0 And ub <= 1 Then
                    xPtemp = x1 + ua * (x2 - x1)
                    yPtemp = y1 + ua * (y2 - y1)
                    'R = Sqr((xPtemp - X1) * (xPtemp - X1) + (yPtemp - Y1) * (yPtemp - Y1))
                    R = (xPtemp - x1) * (xPtemp - x1) + (yPtemp - y1) * (yPtemp - y1)
                    If R < minR And R <> 0 Then
                        If q > CountCurves Then
                            xp = xPtemp
                            yp = yPtemp
                        Else
                            r1 = (xPtemp - x3) * (xPtemp - x3) + (yPtemp - y3) * (yPtemp - y3)
                            R2 = (xPtemp - x4) * (xPtemp - x4) + (yPtemp - y4) * (yPtemp - y4)
                            If r1 < R2 Then
                                xp = x3
                                yp = y3
                            Else
                                xp = x4
                                yp = y4
                            End If
                        End If
                        minR = R
                        OverCrossTest = True
                        qSec = q
                        q2Sec = q2
                    End If
                End If
10
            Next
20
        Next
30
    Next
    If OverCrossTest = True Then
        If TestPointIn(x1, y1, Curves(qSec).SubPath(q2Sec)) = True Then
            SecError = True
        End If
        If GlobalCrossTest(x1, y1, xp, yp, Curves) = True Then
            SecError = True
        End If
    End If
End Function

' -----------------------------------------------------------------------------
' TestPointIn (Function), source lines 1223-1260
' -----------------------------------------------------------------------------
Private Function TestPointIn(ByVal x1 As Single, ByVal y1 As Single, ByRef SubPath As SubPath) As Boolean
    TestPointIn = False
    'Проверяет принадлежность точки к фигуре
    Dim q3 As Long
    Dim x2 As Single
    Dim x3 As Single, y3 As Single
    Dim x4 As Single, y4 As Single
    Dim chis1 As Single, chis2 As Single, znamen As Single
    Dim ua As Single, ub As Single
    x2 = -999999999999999#

    With SubPath
        For q3 = 1 To SubPath.CountNodes - 1
            x3 = .PosX(q3)
            x4 = .PosX(q3 + 1)
            If x3 > x1 And x4 > x1 Then GoTo 10
            
            y3 = .PosY(q3)
            y4 = .PosY(q3 + 1)
            If y3 > y1 And y4 > y1 Then GoTo 10
            If y3 < y1 And y4 < y1 Then GoTo 10
            
            znamen = (y4 - y3) * (x2 - x1)
            chis1 = (x4 - x3) * (y1 - y3) - (y4 - y3) * (x1 - x3)
            chis2 = (x2 - x1) * (y1 - y3)
            If znamen = 0 Then GoTo 10
            
            ua = chis1 / znamen
            ub = chis2 / znamen
    
            If ua >= 0 And ua <= 1 And ub >= 0 And ub <= 1 Then
                'пересекается
                TestPointIn = Not TestPointIn
            End If
10
        Next
    End With
End Function
