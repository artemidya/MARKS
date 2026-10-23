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
