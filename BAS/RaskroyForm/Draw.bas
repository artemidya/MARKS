Private Sub Draw(ByVal matW As Double, ByVal matH As Double, _
                 ByVal plotW As Double, ByVal plotH As Double, _
                 ByVal uF As Double, ByVal lF As Double, _
                 ByVal rF As Double, ByVal dF As Double)
    On Error GoTo 50
    NotNest = False
    If CountResults > 0 Then Command3.Enabled = True
    Command6.Enabled = False
    SelectedResult = CountResults + 1
    Command4.Enabled = False
    Cls True
    If Check2.value Then
        GoTo FigurniyList
        Exit Sub
    End If
    
    'Рисует страницу только для ознакомления
    Dim wpix As Double
    Dim hpix As Double
    Dim ks As Double
    Dim dx As Double
    Dim dy As Double
    Dim oUF As Double, oLF As Double
    Dim oRF As Double, oDF As Double
    
    If plotW = 0 Or plotH = 0 Then NotNest = True
    If matW = 0 Or matH = 0 Then
        NotNest = True
        Exit Sub
    End If
    
    
    If matW / WindowWidth > matH / WindowHeigth Then
        ks = (WindowWidth - 15) / matW
    Else
        ks = (WindowHeigth - 20) / matH
    End If
    
    wpix = ks * matW
    hpix = ks * matH
    
    dx = 6
    dy = 8

    FillRect ShadowBrush, dx + 2, WindowHeigth - hpix - dy + 2, wpix, hpix
    FillRect PaperBrush, dx, WindowHeigth - hpix - dy, wpix, hpix
    
    Dim fieldError As Boolean
    Dim str As String
    If lF + rF > plotW Or lF + rF > matW Then
        fieldError = True
    End If
    If uF + dF > plotH Or uF + dF > matH Then
        fieldError = True
    End If
    
    wpix = ks * plotW
    hpix = ks * plotH
    oUF = ks * uF
    oLF = ks * lF
    oRF = ks * rF
    oDF = ks * dF
    If plotW > 0 And plotH > 0 And fieldError = False Then
        If matH / plotH < 100 Then
            Dim ch As Double, c As Double
            For ch = 1 To matH / plotH
                If matW / plotW < 100 Then
                    For c = 1 To matW / plotW
                        DrawRect PagesPen, (c - 1) * wpix + dx, -(ch - 0) * hpix + WindowHeigth - dy, wpix, hpix
                        DrawRect FieldsPen, (c - 1) * wpix + dx + oLF, -(ch - 0) * hpix + WindowHeigth - dy + oUF, wpix - oLF - oRF, hpix - oUF - oDF
                    Next
                    If matW / plotW > Int(matW / plotW) Then
                        DrawRect PagesPen, (c - 1) * wpix + dx, -(ch - 0) * hpix + WindowHeigth - dy, wpix, hpix
                        DrawRect FieldsPen, (c - 1) * wpix + dx + oLF, -(ch - 0) * hpix + WindowHeigth - dy + oUF, wpix - oLF - oRF, hpix - oUF - oDF
                    End If
                End If
            Next
            If matH / plotH > Int(matH / plotH) Then
                If matW / plotW < 100 Then
                    For c = 1 To matW / plotW
                        DrawRect PagesPen, (c - 1) * wpix + dx, -(ch - 0) * hpix + WindowHeigth - dy, wpix, hpix
                        DrawRect FieldsPen, (c - 1) * wpix + dx + oLF, -(ch - 0) * hpix + WindowHeigth - dy + oUF, wpix - oLF - oRF, hpix - oUF - oDF
                    Next
                    If matW / plotW > Int(matW / plotW) Then
                        DrawRect PagesPen, (c - 1) * wpix + dx, -(ch - 0) * hpix + WindowHeigth - dy, wpix, hpix
                        DrawRect FieldsPen, (c - 1) * wpix + dx + oLF, -(ch - 0) * hpix + WindowHeigth - dy + oUF, wpix - oLF - oRF, hpix - oUF - oDF
                    End If
                End If
            End If
        End If
    End If

    wpix = ks * matW
    hpix = ks * matH
    DrawRect LinePen, dx, WindowHeigth - hpix - dy, wpix, hpix
    
    If lF + rF > plotW Or lF + rF > matW Then
        NotNest = True
        str = localization.ActiveLng.Item("RaskroyForm.TooMuchLeftOrRightField")
        TextOut HDCForm, shiftX + 10, shiftY + WindowHeigth - 30, str, Len(str)
    End If
    If uF + dF > plotH Or uF + dF > matH Then
        NotNest = True
        str = localization.ActiveLng.Item("RaskroyForm.TooHighUpperOrLowerField")
        TextOut HDCForm, shiftX + 10, shiftY + WindowHeigth - 30, str, Len(str)
    End If
    
    Exit Sub
FigurniyList:
    '=========================================================
    '           Здесь обработка для фигурного листа
    '=========================================================

    If AlfaWidth / WindowWidth > AlfaHeight / WindowHeigth Then
        ks = (WindowWidth - 15) / AlfaWidth
    Else
        ks = (WindowHeigth - 20) / AlfaHeight
    End If
    
    'NumberOfMaxFigure
    Dim q As Long, q2 As Long
    Dim sNX As Double, sNY As Double
    Dim eNX As Double, eNY As Double
    Dim SCPx As Double, SCPy As Double
    Dim ECPx As Double, ECPy As Double
    Dim minY As Double, maxY As Double
    Dim minX As Double, maxX As Double
    Dim aa As Long
    Dim sheetCurve As PLCurve
    Dim maxRecurse As Long
    Dim a As Double, a1 As Double, a2 As Double
    Dim da1 As Double, da2 As Double
    minX = 999999999999#
    minY = 999999999999#
    
    sheetCurve.CountPaths = UBound(RCurves(NumberOfMaxFigure).RPaths)
    ReDim sheetCurve.Paths(1 To sheetCurve.CountPaths)
    With RCurves(NumberOfMaxFigure)
        For q = 1 To UBound(.RPaths)
            aa = 1
            sNX = .RPaths(q).RSegments(1).SNPx * RCurves(NumberOfMaxFigure).MaxLenght * ks
            sNY = .RPaths(q).RSegments(1).SNPy * RCurves(NumberOfMaxFigure).MaxLenght * ks
                ReDim sheetCurve.Paths(q).segment(1 To 1)
                sheetCurve.Paths(q).segment(1).Xpoint = sNX
                sheetCurve.Paths(q).segment(1).Ypoint = sNY
            With .RPaths(q)
                For q2 = 1 To UBound(.RSegments)
                    SCPx = .RSegments(q2).SCPx * RCurves(NumberOfMaxFigure).MaxLenght * ks
                    SCPy = .RSegments(q2).SCPy * RCurves(NumberOfMaxFigure).MaxLenght * ks
                    ECPx = .RSegments(q2).ECPx * RCurves(NumberOfMaxFigure).MaxLenght * ks
                    ECPy = .RSegments(q2).ECPy * RCurves(NumberOfMaxFigure).MaxLenght * ks
                    If q2 = UBound(.RSegments) Then
                        eNX = sheetCurve.Paths(q).segment(1).Xpoint
                        eNY = sheetCurve.Paths(q).segment(1).Ypoint
                    Else
                        eNX = .RSegments(q2 + 1).SNPx * RCurves(NumberOfMaxFigure).MaxLenght * ks
                        eNY = .RSegments(q2 + 1).SNPy * RCurves(NumberOfMaxFigure).MaxLenght * ks
                    End If
                        a1 = Angel(sNX, sNY, SCPx, SCPy)
                        a = Angel(SCPx, SCPy, ECPx, ECPy)
                        a2 = Angel(ECPx, ECPy, eNX, eNY)
                        da1 = Abs(DeltaAngle(a, a1))
                        da2 = Abs(DeltaAngle(a2, a))
                        If SCPx = sNX And SCPy = sNY Then da1 = 0
                        If ECPx = eNX And ECPy = eNY Then da2 = 0
                        a = da1 + da2
                        If a < 0.01 Then
                            maxRecurse = 0
                        ElseIf a < 0.05 Then
                            maxRecurse = 1
                        ElseIf a < 0.26 Then
                            maxRecurse = 2
                        ElseIf a < 0.9 Then
                            maxRecurse = 3
                        ElseIf a < 2.2 Then
                            maxRecurse = 4
                        Else
                            maxRecurse = 5
                        End If
                        recursive_bezier sNX, sNY, SCPx, SCPy, ECPx, ECPy, eNX, eNY, 1, sheetCurve.Paths(q), aa, maxRecurse, maxX, minX, maxY, minY
                    sNX = eNX
                    sNY = eNY
                Next
            End With
        Next
    End With
    'Теперь нарисовать лист
    dx = 6
    dy = 8
    For q = 1 To UBound(sheetCurve.Paths)
        With sheetCurve.Paths(q)
            For q2 = 1 To UBound(.segment)
                .segment(q2).Ypoint = -dy + WindowHeigth - .segment(q2).Ypoint + minY
                .segment(q2).Xpoint = dx + .segment(q2).Xpoint - minX
            Next
        End With
    Next
    Dim lines() As LinesString
    Dim y As Long
    Dim y1 As Double, y2 As Double
    Dim x1 As Double, x2 As Double
    ReDim lines(1 To WindowHeigth) As LinesString
    For y = WindowHeigth - (maxY - minY) - dy To WindowHeigth
        For q = 1 To UBound(sheetCurve.Paths)
            With sheetCurve.Paths(q)
                For q2 = 1 To UBound(.segment) - 1
                    y1 = .segment(q2).Ypoint
                    y2 = .segment(q2 + 1).Ypoint
                    If y > y1 And y > y2 Then GoTo nx
                    If y < y1 And y < y2 Then GoTo nx
                    If y2 = y1 Then GoTo nx
                    x1 = .segment(q2).Xpoint
                    x2 = .segment(q2 + 1).Xpoint
                    lines(y).CountX = lines(y).CountX + 1
                    ReDim Preserve lines(y).x(1 To lines(y).CountX)
                    lines(y).x(lines(y).CountX) = (x2 - x1) * (y - y1) / (y2 - y1) + x1
nx:
                Next
            End With
        Next
    Next
    On Error Resume Next
    For q = 1 To UBound(lines)
        With lines(q)
            SortData .x()
        End With
    Next
    For q = 1 To UBound(lines)
        With lines(q)
            For q2 = 1 To .CountX Step 2
                 DrawLine LinesPaperPen, .x(q2), q, .x(q2 + 1), q
                 DrawLine LinesShadowPen, .x(q2) + 2, q + 2, .x(q2 + 1) + 2, q + 2
            Next
        End With
    Next
    For q = 1 To UBound(sheetCurve.Paths)
        With sheetCurve.Paths(q)
            For q2 = 1 To UBound(.segment) - 1
                DrawLine LinePen, .segment(q2).Xpoint, .segment(q2).Ypoint, .segment(q2 + 1).Xpoint, .segment(q2 + 1).Ypoint
            Next
        End With
    Next
50
End Sub
