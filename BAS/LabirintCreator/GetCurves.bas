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
