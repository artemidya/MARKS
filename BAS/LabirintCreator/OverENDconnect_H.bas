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
