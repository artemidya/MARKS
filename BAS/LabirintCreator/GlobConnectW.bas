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
