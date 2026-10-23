Private Function TakeRCurve(curve As VGCore.curve) As RCurve
    Dim q2 As Long, q3 As Long
    Dim minY As Double, maxY As Double
    Dim minX As Double, maxX As Double
    Dim pRline As Boolean
    Dim plusDx As Double, plusDy As Double
    
    TakeRCurve.Width = curve.boundingBox.Width
    TakeRCurve.Height = curve.boundingBox.Height
    Dim c As curve
    
    
    minX = 999999999
    minY = 999999999
    maxX = -999999999
    maxY = -999999999
    If TakeRCurve.Width + TakeRCurve.Height > 0.001 Then
        ReDim TakeRCurve.RPaths(1 To curve.Subpaths.count)
        For q2 = 1 To UBound(TakeRCurve.RPaths)
            With curve.Subpaths(q2)
                'If .Closed = False Then .Closed = True
                ReDim TakeRCurve.RPaths(q2).RSegments(1 To .segments.count)
                pRline = False
                For q3 = 1 To UBound(TakeRCurve.RPaths(q2).RSegments)
                    With .segments(q3)
                        TakeRCurve.RPaths(q2).RSegments(q3).SNPx = .StartNode.PositionX
                        TakeRCurve.RPaths(q2).RSegments(q3).SNPy = .StartNode.PositionY
                        If pRline = True Then
                            TakeRCurve.RPaths(q2).RSegments(q3 - 1).ECPx = TakeRCurve.RPaths(q2).RSegments(q3).SNPx
                            TakeRCurve.RPaths(q2).RSegments(q3 - 1).ECPy = TakeRCurve.RPaths(q2).RSegments(q3).SNPy
                            pRline = False
                        End If
                        With TakeRCurve.RPaths(q2).RSegments(q3)
                            If .SNPx > maxX Then maxX = .SNPx
                            If .SNPx < minX Then minX = .SNPx
                            If .SNPy > maxY Then maxY = .SNPy
                            If .SNPy < minY Then minY = .SNPy
                        End With
                        If .Type = cdrCurveSegment Then
                            TakeRCurve.RPaths(q2).RSegments(q3).SCPx = .StartingControlPointX
                            TakeRCurve.RPaths(q2).RSegments(q3).SCPy = .StartingControlPointY
                            TakeRCurve.RPaths(q2).RSegments(q3).ECPx = .EndingControlPointX
                            TakeRCurve.RPaths(q2).RSegments(q3).ECPy = .EndingControlPointY
                            With TakeRCurve.RPaths(q2).RSegments(q3)
                                If .SCPx > maxX Then maxX = .SCPx
                                If .SCPx < minX Then minX = .SCPx
                                If .SCPy > maxY Then maxY = .SCPy
                                If .SCPy < minY Then minY = .SCPy
                                If .ECPx > maxX Then maxX = .ECPx
                                If .ECPx < minX Then minX = .ECPx
                                If .ECPy > maxY Then maxY = .ECPy
                                If .ECPy < minY Then minY = .ECPy
                            End With
                        Else
                            With TakeRCurve.RPaths(q2).RSegments(q3)
                                .SCPx = .SNPx
                                .SCPy = .SNPy
                                pRline = True
                            End With
                        End If
                    End With
                Next
                    If pRline = True Then
                        TakeRCurve.RPaths(q2).RSegments(q3 - 1).ECPx = TakeRCurve.RPaths(q2).RSegments(1).SNPx
                        TakeRCurve.RPaths(q2).RSegments(q3 - 1).ECPy = TakeRCurve.RPaths(q2).RSegments(1).SNPy
                    End If
            End With
        Next
    Else
        If TakeRCurve.Width + TakeRCurve.Height > 0.001 Then
            maxX = curve.boundingBox.Width + curve.boundingBox.x
            minX = curve.boundingBox.x
            maxY = curve.boundingBox.Height + curve.boundingBox.y
            minY = curve.boundingBox.y
        Else
            maxX = 0.001
            maxY = 0.001
            minX = 0
            minY = 0
        End If
        ReDim TakeRCurve.RPaths(1 To 1)
        ReDim TakeRCurve.RPaths(1).RSegments(1 To 4)
        With TakeRCurve.RPaths(1)
            .RSegments(1).SNPx = minX
            .RSegments(1).SNPy = minY
            .RSegments(1).SCPx = minX
            .RSegments(1).SCPy = minY
            .RSegments(1).ECPx = maxX
            .RSegments(1).ECPy = minY
            
            .RSegments(2).SNPx = maxX
            .RSegments(2).SNPy = minY
            .RSegments(2).SCPx = maxX
            .RSegments(2).SCPy = minY
            .RSegments(2).ECPx = maxX
            .RSegments(2).ECPy = maxY
            
            .RSegments(3).SNPx = maxX
            .RSegments(3).SNPy = maxY
            .RSegments(3).SCPx = maxX
            .RSegments(3).SCPy = maxY
            .RSegments(3).ECPx = minX
            .RSegments(3).ECPy = maxY
            
            .RSegments(4).SNPx = minX
            .RSegments(4).SNPy = maxY
            .RSegments(4).SCPx = minX
            .RSegments(4).SCPy = maxY
            .RSegments(4).ECPx = minX
            .RSegments(4).ECPy = minY
        End With
    End If
    '“епрь надо перевести в относительную систему координат
    For q2 = 1 To UBound(TakeRCurve.RPaths)
        With TakeRCurve.RPaths(q2)
            For q3 = 1 To UBound(.RSegments())
                With .RSegments(q3)
                    .ECPx = SupRound(.ECPx - minX)
                    .ECPy = SupRound(.ECPy - minY)
                    .SCPx = SupRound(.SCPx - minX)
                    .SCPy = SupRound(.SCPy - minY)
                    .SNPx = SupRound(.SNPx - minX)
                    .SNPy = SupRound(.SNPy - minY)
                End With
            Next
        End With
    Next
    
    maxX = SupRound(maxX - minX)
    minX = 0
    maxY = SupRound(maxY - minY)
    minY = 0
    If (maxX) > (maxY) Then
        TakeRCurve.MaxLenght = maxX
        plusDx = 0
        plusDy = SupRound((TakeRCurve.MaxLenght - maxY) / 2)
    Else
        TakeRCurve.MaxLenght = maxY
        plusDx = SupRound((TakeRCurve.MaxLenght - maxX) / 2)
        plusDy = 0
    End If

    If TakeRCurve.MaxLenght = 0 Then TakeRCurve.MaxLenght = 1
    TakeRCurve.MaxLenght = TakeRCurve.MaxLenght / 65530
    '===================================================================================
    'ѕерестроить числа
        For q2 = 1 To UBound(TakeRCurve.RPaths)
            For q3 = 1 To UBound(TakeRCurve.RPaths(q2).RSegments)
                With TakeRCurve.RPaths(q2).RSegments(q3)
                    .SNPx = (.SNPx - minX + plusDx) / TakeRCurve.MaxLenght
                    .SNPy = (.SNPy - minY + plusDy) / TakeRCurve.MaxLenght
                    .SCPx = (.SCPx - minX + plusDx) / TakeRCurve.MaxLenght
                    .SCPy = (.SCPy - minY + plusDy) / TakeRCurve.MaxLenght
                    .ECPx = (.ECPx - minX + plusDx) / TakeRCurve.MaxLenght
                    .ECPy = (.ECPy - minY + plusDy) / TakeRCurve.MaxLenght
                End With
            Next
        Next
End Function
