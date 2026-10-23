' =============================================================================
' Group: 03_Geometry_and_Curves
' Source module: CurveTester.bas
' Procedures: 4
' =============================================================================

' -----------------------------------------------------------------------------
' IntersectionSearch (Sub), source lines 23-65
' -----------------------------------------------------------------------------
Public Sub IntersectionSearch()
    On Error GoTo 30
    UserData.Load
    Dim lastUnit As cdrUnit
    lastUnit = Application.ActiveDocument.Unit
    Application.ActiveDocument.Unit = cdrInch
    
    If Application.ActiveSelection.shapes.count = 0 Then Exit Sub
    
    Application.Optimization = True
    Application.ActiveDocument.BeginCommandGroup localization.ActiveLng.Item("CurveTester.BeginCommandGroup")

    Application.ActiveSelection.ConvertToCurves
    Application.ActiveSelection.UngroupAllEx
    
    
    Dim Curves() As curve
    Dim CountCurves As Long
    
    GetCurves Curves(), CountCurves
    
    ActiveDocument.Undo 1
    
    
    Dim countCrossPoints As Long
    Dim x() As Double, y() As Double
    countCrossPoints = CrossPoints(x(), y(), Curves(), CountCurves)
    

    If countCrossPoints > 0 Then
        If MsgBox(localization.ActiveLng.Item("CurveTester.MsgBox.P1") & " " & countCrossPoints & Chr(10) & Chr(13) & localization.ActiveLng.Item("CurveTester.MsgBox.P2"), vbYesNo, "PlotCalc") = vbYes Then
            DrivePoints x(), y(), countCrossPoints
        End If
    Else
        MsgBox localization.ActiveLng.Item("CurveTester.MsgBox2"), vbInformation, "PlotCalc"
    End If
    Application.ActiveDocument.EndCommandGroup

30
    Application.ActiveDocument.Unit = lastUnit
    Application.Optimization = False
    Application.Refresh
End Sub

' -----------------------------------------------------------------------------
' GetCurves (Sub), source lines 67-172
' -----------------------------------------------------------------------------
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

' -----------------------------------------------------------------------------
' CrossPoints (Function), source lines 173-252
' -----------------------------------------------------------------------------
Private Function CrossPoints(ByRef XS() As Double, ByRef YS() As Double, ByRef Curves() As curve, ByRef CountCurves As Long) As Long
    Dim q1 As Long, q2 As Long, q3 As Long
    Dim a1 As Long, a2 As Long, a3 As Long
    Dim x1 As Double, x2 As Double
    Dim y1 As Double, y2 As Double
    Dim x3 As Double, x4 As Double
    Dim y3 As Double, y4 As Double
    
    Dim ua As Double, ub As Double
    Dim znamen As Double, chis1 As Double, chis2 As Double
    
    For q1 = 1 To CountCurves - 1
        With Curves(q1)
            For q2 = 1 To .CountSubPath
                For q3 = 1 To .SubPath(q2).CountNodes - 1
                    With .SubPath(q2)
                        x1 = .PosX(q3)
                        x2 = .PosX(q3 + 1)
                        y1 = .PosY(q3)
                        y2 = .PosY(q3 + 1)
                    End With
                        '================================
        
                    For a1 = q1 + 1 To CountCurves
                        With Curves(a1)
                            x3 = .PositionX
                            x4 = .PositionX + .With
                            If x3 > x1 And x4 > x1 And x3 > x2 And x4 > x2 Then GoTo 30
                            If x3 < x1 And x4 < x1 And x3 < x2 And x4 < x2 Then GoTo 30
        
                            y3 = .PositionY - .Height
                            y4 = .PositionY
                            If y3 > y1 And y4 > y1 And y3 > y2 And y4 > y2 Then GoTo 30
                            If y3 < y1 And y4 < y1 And y3 < y2 And y4 < y2 Then GoTo 30
                            For a2 = 1 To .CountSubPath
                                With .SubPath(a2)
                                    x3 = .PositionX
                                    x4 = .PositionX + .With
                                    If x3 > x1 And x4 > x1 And x3 > x2 And x4 > x2 Then GoTo 20
                                    If x3 < x1 And x4 < x1 And x3 < x2 And x4 < x2 Then GoTo 20
                                    y3 = .PositionY - .Height
                                    y4 = .PositionY
                                    If y3 > y1 And y4 > y1 And y3 > y2 And y4 > y2 Then GoTo 20
                                    If y3 < y1 And y4 < y1 And y3 < y2 And y4 < y2 Then GoTo 20
                                    For a3 = 1 To .CountNodes - 1
                                        x3 = .PosX(a3)
                                        x4 = .PosX(a3 + 1)
                                        If x3 > x1 And x4 > x1 And x3 > x2 And x4 > x2 Then GoTo 10
                                        If x3 < x1 And x4 < x1 And x3 < x2 And x4 < x2 Then GoTo 10
                                        y3 = .PosY(a3)
                                        y4 = .PosY(a3 + 1)
                                        If y3 > y1 And y4 > y1 And y3 > y2 And y4 > y2 Then GoTo 10
                                        If y3 < y1 And y4 < y1 And y3 < y2 And y4 < y2 Then GoTo 10
                                        '---Проверка на пересечение---
                                        znamen = (y4 - y3) * (x2 - x1) - (x4 - x3) * (y2 - y1)
                                        If znamen = 0 Then GoTo 10
                                        chis1 = (x4 - x3) * (y1 - y3) - (y4 - y3) * (x1 - x3)
                                        chis2 = (x2 - x1) * (y1 - y3) - (y2 - y1) * (x1 - x3)
                                        ua = chis1 / znamen
                                        ub = chis2 / znamen
                                        If ua > 0 And ua <= 1 And ub >= 0 And ub < 1 Then
                                            CrossPoints = CrossPoints + 1
                                            ReDim Preserve XS(CrossPoints)
                                            ReDim Preserve YS(CrossPoints)
                                            XS(CrossPoints) = x1 + ua * (x2 - x1)
                                            YS(CrossPoints) = y1 + ua * (y2 - y1)
                                        End If
10
                                    Next
20
                                End With
                            Next
30
                        End With
                    Next
                Next
            Next
        End With
    Next
End Function

' -----------------------------------------------------------------------------
' DrivePoints (Sub), source lines 253-270
' -----------------------------------------------------------------------------
Private Sub DrivePoints(ByRef XS() As Double, ByRef YS() As Double, ByVal count As Long)
    Dim q As Long
    Dim n As Double
    Dim s1 As shape
    Dim crv As VGCore.curve
    
    Set crv = Application.CreateCurve(ActiveDocument)
    For q = 1 To count
        n = 0.2
        crv.CreateSubPath(XS(q) - n, YS(q) - n).AppendLineSegment XS(q) + n, YS(q) + n
        crv.CreateSubPath(XS(q) - n, YS(q) + n).AppendLineSegment XS(q) + n, YS(q) - n
    Next
    
    Set s1 = Application.ActiveLayer.CreateCurve(crv) '.Outline.SetProperties(Color:=CreateCMYKColor(0, 60, 100, 0))
    s1.Outline.SetProperties Color:=CreateCMYKColor(0, 100, 100, 0)
    's1.Outline.SetProperties 0.04222, OutlineStyles(0), CreateCMYKColor(0, 100, 100, 0), ArrowHeads(0), ArrowHeads(0), False, False, cdrOutlineButtLineCaps, cdrOutlineMiterLineJoin, 0#, 100

End Sub
