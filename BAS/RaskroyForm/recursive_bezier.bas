Private Sub recursive_bezier(x1 As Double, y1 As Double, x2 As Double, y2 As Double, x3 As Double, y3 As Double, x4 As Double, y4 As Double, R As Double, ByRef path As path, ByRef aa As Long, ByRef maxRecurse As Long, _
                            ByRef maxX As Double, ByRef minX As Double, ByRef maxY As Double, ByRef minY As Double)
   ' Вычислить все средние точек отрезков
    Dim x12 As Double, y12 As Double
    Dim x23 As Double, y23 As Double
    Dim x34 As Double, y34 As Double
    Dim x123 As Double, y123 As Double
    Dim x234 As Double, y234 As Double
    Dim x1234 As Double, y1234 As Double
    
    x12 = (x1 + x2) / 2
    y12 = (y1 + y2) / 2
    x23 = (x2 + x3) / 2
    y23 = (y2 + y3) / 2
    x34 = (x3 + x4) / 2
    y34 = (y3 + y4) / 2
    x123 = (x12 + x23) / 2
    y123 = (y12 + y23) / 2
    x234 = (x23 + x34) / 2
    y234 = (y23 + y34) / 2
    x1234 = (x123 + x234) / 2
    y1234 = (y123 + y234) / 2
    If R > maxRecurse Then
        ' Нарисовать отрезок и закончить
        aa = aa + 1
        ReDim Preserve path.segment(1 To aa)
        path.segment(aa).Xpoint = x4
        path.segment(aa).Ypoint = y4
        If x4 > maxX Then maxX = x4
        If x4 < minX Then minX = x4
        If y4 > maxY Then maxY = y4
        If y4 < minY Then minY = y4
    Else
        ' Продолжить деление
        R = R + 1
        recursive_bezier x1, y1, x12, y12, x123, y123, x1234, y1234, R, path, aa, maxRecurse, maxX, minX, maxY, minY
        recursive_bezier x1234, y1234, x234, y234, x34, y34, x4, y4, R, path, aa, maxRecurse, maxX, minX, maxY, minY
        R = R - 1
   End If
End Sub
