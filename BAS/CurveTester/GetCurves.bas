Private Sub GetCurves(ByRef Curves() As curve, ByRef CountCurves As Long)
    Dim q As Long, q2 As Long, q3 As Long, q4 As Long
    Dim n As Long, n2 As Long, n3 As Long, n4 As Long
    Dim a As Long, s As Long
    Dim k As Long, stp As Double, t As Double
    Dim p0x As Double, p0y As Double
    Dim p1x As Double, p1y As Double
    Dim p2x As Double, p2y As Double
    Dim p3x As Double, p3y As Double
    Dim leng As Double
    
    Dim selObj As shape
    Set selObj = ActiveSelection
    
    
    
    n = selObj.shapes.count
    CountCurves = n
    ReDim Curves(n)
    '--------------------
    
    For q = 1 To n
        With selObj.shapes(q)
            n2 = .curve.Subpaths.count
            Curves(q).PositionX = .PositionX
            Curves(q).PositionY = .PositionY
            Curves(q).With = .SizeWidth
            Curves(q).Height = .SizeHeight
            Curves(q).CountSubPath = n2
            ReDim Curves(q).SubPath(n2)
            For q2 = 1 To n2
                With .curve.Subpaths(q2)
                    Curves(q).SubPath(q2).PositionX = .PositionX
                    Curves(q).SubPath(q2).PositionY = .PositionY
                    Curves(q).SubPath(q2).With = .SizeWidth
                    Curves(q).SubPath(q2).Height = .SizeHeight
                    Curves(q).SubPath(q2).CountNodes = .Nodes.count
                    Curves(q).SubPath(q2).IsClosed = .closed
                    'If .Closed = False Then curves(q).SubPath(q2).IsClosed = False Else curves(q).SubPath(q2).IsClosed = True
                End With
            Next
        End With
    Next
    
    '--------------------
    For q = 1 To CountCurves
        For q2 = 1 To Curves(q).CountSubPath
            n3 = Curves(q).SubPath(q2).CountNodes
            If Curves(q).SubPath(q2).IsClosed = False Then n3 = n3 - 1
            a = 0
            p0x = selObj.shapes(q).curve.Subpaths(q2).segments(1).StartNode.PositionX
            p0y = selObj.shapes(q).curve.Subpaths(q2).segments(1).StartNode.PositionY
            For q3 = 1 To n3
                With selObj.shapes(q).curve.Subpaths(q2).segments(q3)
                    'p0x = .StartNode.PositionX
                    'p0y = .StartNode.PositionY
                    p3x = .EndNode.PositionX
                    p3y = .EndNode.PositionY
                    If .Type = cdrLineSegment Then
                        a = a + 1
                        ReDim Preserve Curves(q).SubPath(q2).PosX(a)
                        ReDim Preserve Curves(q).SubPath(q2).PosY(a)
                        Curves(q).SubPath(q2).PosX(a) = p0x
                        Curves(q).SubPath(q2).PosY(a) = p0y
                    Else
                        p1x = .StartingControlPointX
                        p1y = .StartingControlPointY
                        p2x = .EndingControlPointX
                        p2y = .EndingControlPointY
                        
                        leng = .Length
                        If leng = 0 Then leng = 0.000000000001
                        k = Int(1 / (0.196850392007875 / leng)) + 1
                        stp = 1 / k
                        a = a + 1
                        With Curves(q).SubPath(q2)
                            ReDim Preserve .PosX(a + k)
                            ReDim Preserve .PosY(a + k)
                            .PosX(a) = p0x
                            .PosY(a) = p0y
                            t = 0
                            For s = 1 To k - 1
                                t = t + stp
                                a = a + 1
                                .PosX(a) = (1 - t) ^ 3 * p0x + 3 * (1 - t) ^ 2 * t * p1x + 3 * (1 - t) * t ^ 2 * p2x + t ^ 3 * p3x
                                .PosY(a) = (1 - t) ^ 3 * p0y + 3 * (1 - t) ^ 2 * t * p1y + 3 * (1 - t) * t ^ 2 * p2y + t ^ 3 * p3y
                            Next
                        End With
                    End If
                End With
                p0x = p3x
                p0y = p3y
            Next
            
            a = a + 1
            With Curves(q).SubPath(q2)
                ReDim Preserve .PosX(a)
                ReDim Preserve .PosY(a)
                .PosX(a) = p3x
                .PosY(a) = p3y
                .CountNodes = a
            End With
        
        Next
    Next
End Sub
