' =============================================================================
' Group: 03_Geometry_and_Curves
' Source module: ReGroup.bas
' Procedures: 20
' =============================================================================

' -----------------------------------------------------------------------------
' StartReGroup (Sub), source lines 9-95
' -----------------------------------------------------------------------------
Public Sub StartReGroup()
    If Processing Then Exit Sub
    UserData.Load
    If VGCore.ActiveSelectionRange.count > 0 Then
        On Error GoTo er:
        Dim RegData1 As New RegData

        Processing = True
        ProgressFormIsShowed = False
        Application.Optimization = True
        Application.ActiveDocument.PreserveSelection = True
        Application.ActiveDocument.BeginCommandGroup localization.ActiveLng.Item("ReGroup.Command")
        
        Dim sr As New VGCore.ShapeRange
        sr.AddRange VGCore.ActiveSelectionRange.UngroupAllEx
        
        StartTime = Timer
        LastRefreshTime = Timer
        
        If RegData1.ConnectingCurves Then
            '1. Connecting
            ProgressForm.StageName = localization.ActiveLng.Item("ReGroup.ConnectingLines")
            ProgressForm.EnableSkip = True
            ConnectingCurves sr
            If ProgressForm.Break Then GoTo er
            ProgressForm.EnableSkip = False
            ProgressForm.Percent = 1
    
            '2. Connecting subpaths
            ProgressForm.StageName = localization.ActiveLng.Item("ReGroup.ConnectingSubpaths")
            SelfyConnectingShapes sr
        End If
        
        '3. Sorting
        ProgressForm.StageName = localization.ActiveLng.Item("ReGroup.Sorting")
        sr.Sort " @shape1.Width * @shape1.Height > @shape2.Width * @shape2.Height"
        
        '4. Add to ShapesArrey
        Dim ShapesArray() As VGCore.shape
        ReDim ShapesArray(sr.count - 1)
        Dim q As Long
        For q = 1 To sr.count
            Set ShapesArray(q - 1) = sr(q)
        Next
        
        '5. ReGrouping
        ProgressForm.StageName = localization.ActiveLng.Item("ReGroup.Command")
        ProgressForm.Percent = 0
        StartR ShapesArray()
        If ProgressForm.Break Then GoTo er
        ProgressForm.Percent = 1
        
        
        Set sr = New VGCore.ShapeRange
        For q = 0 To UBound(ShapesArray())
            sr.Add ShapesArray(q)
        Next
    
        If RegData1.ReduceCurves Then
            '6. Reduce
            ProgressForm.EnableSkip = True
            ProgressForm.StageName = localization.ActiveLng.Item("ReGroup.ReduceNodes")
            CountShapes = 0
            ReadyCountShapes = 0
            GetCountCurves sr.shapes
            Reduce sr.shapes
            ProgressForm.EnableSkip = False
        End If
        
        Application.ActiveDocument.ClearSelection
        sr.All.AddToSelection
        
er:
        ProgressForm.MbClosing = True
        Application.ActiveDocument.EndCommandGroup
        On Error Resume Next
        If ProgressForm.Break Then
            Application.ActiveDocument.Undo
        End If
        Unload ProgressForm
        
        Application.Optimization = False
        Application.ActiveDocument.PreserveSelection = False
        Processing = False
        Application.Refresh
    End If
End Sub

' -----------------------------------------------------------------------------
' GetCountCurves (Sub), source lines 96-107
' -----------------------------------------------------------------------------
Private Sub GetCountCurves(shapes As VGCore.shapes)
    Dim q As Long
    For q = 1 To shapes.count
        With shapes(q)
            If .Type = cdrCurveShape Then
               CountShapes = CountShapes + 1
            ElseIf .Type = cdrGroupShape Then
               GetCountCurves .shapes
            End If
        End With
    Next
End Sub

' -----------------------------------------------------------------------------
' Reduce (Sub), source lines 108-141
' -----------------------------------------------------------------------------
Private Sub Reduce(shapes As VGCore.shapes)
    Dim q As Long
    Dim txt As String
    txt = localization.ActiveLng.Item("ReGroup.ReduceNodes")
    For q = 1 To shapes.count
        With shapes(q)
            '------Progress-Bar-------
            If ProgressFormIsShowed Then
                If Timer - LastRefreshTime > 0.1 Then
                    ProgressForm.Percent = ReadyCountShapes / CountShapes
                    LastRefreshTime = Timer
                    ProgressForm.StageName = txt & " (" & ReadyCountShapes & " / " & CountShapes & ")"
                    DoEvents
                    If ProgressForm.Break Or ProgressForm.Skip Then
                        Exit Sub
                    End If
                End If
            Else
                If Timer - StartTime > 0.5 Then
                    ProgressFormIsShowed = True
                    ProgressForm.Show
                End If
            End If
            '------Progress-Bar-------
        
            If .Type = cdrCurveShape Then
                ReadyCountShapes = ReadyCountShapes + 1
                .curve.AutoReduceNodes 0
            ElseIf .Type = cdrGroupShape Then
                Reduce .shapes
            End If
        End With
    Next
End Sub

' -----------------------------------------------------------------------------
' ConnectingCurves (Sub), source lines 309-390
' -----------------------------------------------------------------------------
Private Sub ConnectingCurves(sr As VGCore.ShapeRange)
    Dim q1 As Long, q2 As Long
    
    Dim left As Double, right As Double
    Dim top As Double, bottom As Double
    Dim crv As VGCore.curve
    Dim count As Long
    Dim txt As String
    txt = localization.ActiveLng.Item("ReGroup.ConnectingLines")
    
    q1 = 1
    With sr
        count = .count
        Do While q1 < count
            If sr(q1).Type = cdrCurveShape Then
                If OneOfSubpathsIsOpened(sr(q1).curve) Then
                    With sr(q1).boundingBox
                        left = .left - 0.00001
                        right = .right + 0.00001
                        top = .top + 0.00001
                        bottom = .bottom - 0.00001
                    End With
                    
                    '------Progress-Bar-------
                    If ProgressFormIsShowed Then
                        If Timer - LastRefreshTime > 0.1 Then
                            ProgressForm.Percent = q1 / count
                            LastRefreshTime = Timer
                            ProgressForm.StageName = txt & " (" & q1 & " / " & count & ")"
                            DoEvents
                            If ProgressForm.Break Or ProgressForm.Skip Then Exit Sub
                        End If
                    Else
                        If Timer - StartTime > 0.5 Then
                            ProgressFormIsShowed = True
                            ProgressForm.Show
                        End If
                    End If
                    '------Progress-Bar-------
                
                    q2 = q1 + 1
                    Do While q2 <= count
                        With sr(q2)
                            If .Type = cdrCurveShape Then
                                With .boundingBox
                                    If .right >= left Then
                                        If .left <= right Then
                                            If .top >= bottom Then
                                                If .bottom <= top Then
                                                    Set crv = sr(q2).curve
                                                    If ConnectingCurveToCurve(sr(q1).curve, crv) Then
                                                        
                                                        If crv.segments.count = 0 Then
                                                            sr.Remove q2
                                                            count = sr.count
                                                        End If

                                                        With sr(q1).boundingBox
                                                            left = .left - 0.00001
                                                            right = .right + 0.00001
                                                            top = .top + 0.00001
                                                            bottom = .bottom - 0.00001
                                                        End With
                                                        q2 = q1 + 1
                                                        GoTo nx:
                                                    End If
                                                End If
                                            End If
                                        End If
                                    End If
                                End With
                            End If
                        End With
                        q2 = q2 + 1
nx:
                    Loop
                End If
            End If
            q1 = q1 + 1
        Loop
    End With
End Sub

' -----------------------------------------------------------------------------
' ConnectingCurveToCurve (Function), source lines 391-403
' -----------------------------------------------------------------------------
Private Function ConnectingCurveToCurve(crv1 As VGCore.curve, crv2 As VGCore.curve) As Boolean
    Dim q As Long
    With crv1
        For q = 1 To .Subpaths.count
            If Not .Subpaths(q).closed Then
                If ConnectingSubpathToCurve(.Subpaths(q), crv2) Then
                    ConnectingCurveToCurve = True
                    If crv2.segments.count = 0 Then Exit Function
                End If
            End If
        Next
    End With
End Function

' -----------------------------------------------------------------------------
' ConnectingSubpathToCurve (Function), source lines 404-423
' -----------------------------------------------------------------------------
Private Function ConnectingSubpathToCurve(SubPath As VGCore.SubPath, crv As VGCore.curve) As Boolean
    Dim q As Long
    With crv
1:
        For q = 1 To .Subpaths.count
            If Not .Subpaths(q).closed Then
                If .Subpaths(q).segments.count > 0 Then
                    If ConnectTwoSubpathsFromDifferentCurves(SubPath, .Subpaths(q)) Then
                        ConnectingSubpathToCurve = True
                        If .segments.count = 0 Then
                            Exit Function
                        Else
                            GoTo 1
                        End If
                    End If
                End If
            End If
        Next
    End With
End Function

' -----------------------------------------------------------------------------
' ConnectTwoSubpathsFromDifferentCurves (Function), source lines 424-464
' -----------------------------------------------------------------------------
Private Function ConnectTwoSubpathsFromDifferentCurves(SubPath1 As VGCore.SubPath, SubPath2 As VGCore.SubPath) As Boolean
    On Error GoTo 10
    If SubPath1.EndNode.PositionX = SubPath2.StartNode.PositionX Then
        If SubPath1.EndNode.PositionY = SubPath2.StartNode.PositionY Then
            ConnectTwoSubpathsFromDifferentCurves = True
            SubPath1.Parent.AppendCurve SubPath2.Extract.curve
            SubPath2.Delete
            SubPath1.EndNode.ConnectWith SubPath1.Parent.Subpaths.Last.StartNode
            Exit Function
        End If
    End If
'    If SubPath1.EndNode.PositionX = SubPath2.EndNode.PositionX Then
'        If SubPath1.EndNode.PositionY = SubPath2.EndNode.PositionY Then
'            ConnectTwoSubpathsFromDifferentCurves = True
'            SubPath1.Parent.AppendCurve SubPath2.Extract.curve
'            SubPath2.Delete
'            SubPath1.EndNode.ConnectWith SubPath1.Parent.Subpaths.Last.EndNode
'            Exit Function
'        End If
'    End If
'    If SubPath1.StartNode.PositionX = SubPath2.StartNode.PositionX Then
'        If SubPath1.StartNode.PositionY = SubPath2.StartNode.PositionY Then
'             ConnectTwoSubpathsFromDifferentCurves = True
'            SubPath1.Parent.AppendCurve SubPath2.Extract.curve
'            SubPath2.Delete
'            SubPath1.StartNode.ConnectWith SubPath1.Parent.Subpaths.Last.StartNode
'            Exit Function
'        End If
'    End If
    If SubPath1.StartNode.PositionX = SubPath2.EndNode.PositionX Then
        If SubPath1.StartNode.PositionY = SubPath2.EndNode.PositionY Then
            ConnectTwoSubpathsFromDifferentCurves = True
            SubPath1.Parent.AppendCurve SubPath2.Extract.curve
            SubPath2.Delete
            SubPath1.StartNode.ConnectWith SubPath1.Parent.Subpaths.Last.EndNode
            Exit Function
        End If
    End If
  
10
End Function

' -----------------------------------------------------------------------------
' SelfyConnectingShapes (Sub), source lines 465-470
' -----------------------------------------------------------------------------
Private Sub SelfyConnectingShapes(sr As VGCore.ShapeRange)
    Dim q As Long
    For q = 1 To sr.count
        If sr(q).Type = cdrCurveShape Then SelfyConnecting sr(q).curve
    Next
End Sub

' -----------------------------------------------------------------------------
' SelfyConnecting (Sub), source lines 471-499
' -----------------------------------------------------------------------------
Private Sub SelfyConnecting(crv As VGCore.curve)
    Dim q As Long, q2 As Long
    With crv
        q = 1
        Do While q <= .Subpaths.count
            If Not .Subpaths(q).closed Then
                q2 = q + 1
                Do While q2 <= .Subpaths.count
                    If Not .Subpaths(q2).closed Then
                        If ConnectTwoSubpaths(.Subpaths(q), .Subpaths(q2)) Then
                            q2 = q + 1
                            GoTo nx:
                        End If
                    End If
                    q2 = q2 + 1
nx:
                Loop
                If .Subpaths(q).Nodes.Last.PositionX = .Subpaths(q).Nodes.First.PositionX Then
                    If .Subpaths(q).Nodes.Last.PositionY = .Subpaths(q).Nodes.First.PositionY Then
                        If .Subpaths(q).Nodes.count > 1 Then
                            .Subpaths(q).closed = True
                        End If
                    End If
                End If
            End If
            q = q + 1
        Loop
    End With
End Sub

' -----------------------------------------------------------------------------
' ConnectTwoSubpaths (Function), source lines 501-530
' -----------------------------------------------------------------------------
Private Function ConnectTwoSubpaths(SubPath1 As VGCore.SubPath, SubPath2 As VGCore.SubPath) As Boolean
    If SubPath1.EndNode.PositionX = SubPath2.StartNode.PositionX Then
        If SubPath1.EndNode.PositionY = SubPath2.StartNode.PositionY Then
            SubPath1.EndNode.ConnectWith SubPath2.StartNode
            ConnectTwoSubpaths = True
            Exit Function
        End If
    End If
    If SubPath1.EndNode.PositionX = SubPath2.EndNode.PositionX Then
        If SubPath1.EndNode.PositionY = SubPath2.EndNode.PositionY Then
            SubPath1.EndNode.ConnectWith SubPath2.EndNode
            ConnectTwoSubpaths = True
            Exit Function
        End If
    End If
    If SubPath1.StartNode.PositionX = SubPath2.StartNode.PositionX Then
        If SubPath1.StartNode.PositionY = SubPath2.StartNode.PositionY Then
            SubPath1.StartNode.ConnectWith SubPath2.StartNode
            ConnectTwoSubpaths = True
            Exit Function
        End If
    End If
    If SubPath1.StartNode.PositionX = SubPath2.EndNode.PositionX Then
        If SubPath1.StartNode.PositionY = SubPath2.EndNode.PositionY Then
            SubPath1.StartNode.ConnectWith SubPath2.EndNode
            ConnectTwoSubpaths = True
            Exit Function
        End If
    End If
End Function

' -----------------------------------------------------------------------------
' StartR (Sub), source lines 531-564
' -----------------------------------------------------------------------------
Private Sub StartR(ShapesArray() As VGCore.shape)
    Dim q As Long
    Dim count As Long

    Dim txt As String
    txt = localization.ActiveLng.Item("ReGroup.Command")
    
    q = UBound(ShapesArray()) - 1
    count = q + 2

    Do While q > -1
        If OneOfSubpathsIsClosed(ShapesArray(q)) Then
            '------Progress-Bar-------
            If ProgressFormIsShowed Then
                If Timer - LastRefreshTime > 0.1 Then
                    ProgressForm.Percent = (count - q) / count
                    LastRefreshTime = Timer
                    ProgressForm.StageName = txt & " (" & count - q & " / " & count & ") "
                    DoEvents
                    If ProgressForm.Break Then Exit Sub
                End If
            Else
                If Timer - StartTime > 0.5 Then
                    ProgressFormIsShowed = True
                    ProgressForm.Show
                End If
            End If
            '------Progress-Bar-------
        
            ToGroup ShapesArray(), q
        End If
        q = q - 1
    Loop
End Sub

' -----------------------------------------------------------------------------
' ToGroup (Sub), source lines 565-608
' -----------------------------------------------------------------------------
Private Sub ToGroup(ShapesArray() As VGCore.shape, StartFromQ As Long)
    Dim q As Long
    Dim x As Double, y As Double
    
    Dim ForGropingSR As New VGCore.ShapeRange
    
    Dim ControlCurve As VGCore.curve
    Set ControlCurve = CreatePolylineCurveFromShape(ShapesArray(StartFromQ))
       
    Dim left As Double, right As Double
    Dim top As Double, bottom As Double
    
    With ControlCurve.boundingBox
        left = .left
        right = .right
        top = .top
        bottom = .bottom
    End With
    
    q = StartFromQ + 1
    Do While q <= UBound(ShapesArray())
        With ShapesArray(q).boundingBox
            x = (.right + .left) / 2
            If x < right And x > left Then
                y = (.top + .bottom) / 2
                If y < top And y > bottom Then
                    If IsShapeOnPolyLineCurve(ControlCurve, x, y, ShapesArray(q)) Then
'                       inside shape!
                        ForGropingSR.Add ShapesArray(q)
                        DeleteShapeFromShapesArray ShapesArray(), q
                        q = q - 1
                    End If
                End If
            End If
        End With
        q = q + 1
    Loop
    
    If ForGropingSR.count > 0 Then
        ForGropingSR.Add ShapesArray(StartFromQ)
        Set ShapesArray(StartFromQ) = ForGropingSR.Group
    End If
    
End Sub

' -----------------------------------------------------------------------------
' OneOfSubpathsIsClosed (Function), source lines 609-625
' -----------------------------------------------------------------------------
Private Function OneOfSubpathsIsClosed(shape As VGCore.shape) As Boolean
    With shape
        If .Type = cdrCurveShape Then
            Dim q As Long
            With .curve
                For q = 1 To .Subpaths.count
                    If .Subpaths(q).closed Then
                        OneOfSubpathsIsClosed = True
                        Exit For
                    End If
                Next
            End With
        Else
            OneOfSubpathsIsClosed = True
        End If
    End With
End Function

' -----------------------------------------------------------------------------
' OneOfSubpathsIsOpened (Function), source lines 626-636
' -----------------------------------------------------------------------------
Private Function OneOfSubpathsIsOpened(curve As VGCore.curve) As Boolean
    Dim q As Long
    With curve
        For q = 1 To .Subpaths.count
            If Not .Subpaths(q).closed Then
                OneOfSubpathsIsOpened = True
                Exit For
            End If
        Next
    End With
End Function

' -----------------------------------------------------------------------------
' CreatePolylineCurveFromShape (Function), source lines 638-649
' -----------------------------------------------------------------------------
Private Function CreatePolylineCurveFromShape(shape As VGCore.shape) As VGCore.curve
    With shape
        If .Type = cdrCurveShape Then
            Set CreatePolylineCurveFromShape = CreatePolylineCurveFromCurve(shape.curve)
        ElseIf .Type = cdrEllipseShape Then
            Set CreatePolylineCurveFromShape = CreatePolylineCurveFromCurve(shape.DisplayCurve)
        Else
            Set CreatePolylineCurveFromShape = Application.CreateCurve
            CreatePolylineCurveFromShape.AppendSubpathRectangle .LeftX, .TopY, .RightX, .BottomY
        End If
    End With
End Function

' -----------------------------------------------------------------------------
' CreatePolylineCurveFromCurve (Function), source lines 650-681
' -----------------------------------------------------------------------------
Private Function CreatePolylineCurveFromCurve(crv As VGCore.curve) As VGCore.curve
    Set CreatePolylineCurveFromCurve = Application.CreateCurve
    Dim q As Long, q2 As Long, q3 As Long
    Dim sp As VGCore.SubPath
    Dim plc As VGCore.curve
    
    With crv
        For q = 1 To .Subpaths.count
            With .Subpaths(q)
                Set sp = CreatePolylineCurveFromCurve.CreateSubPath(.StartNode.PositionX, .StartNode.PositionY)
                
                For q2 = 1 To .segments.count
                    With .segments(q2)
                        If .Type = cdrCurveSegment Then
                            Set plc = .GetPolyline
                            With plc
                                For q3 = 1 To plc.segments.count
                                    With .segments(q3)
                                        sp.AppendLineSegment .EndNode.PositionX, .EndNode.PositionY
                                    End With
                                Next
                            End With
                        Else
                            sp.AppendLineSegment .EndNode.PositionX, .EndNode.PositionY
                        End If
                    End With
                Next
                sp.closed = .closed
            End With
        Next
    End With
End Function

' -----------------------------------------------------------------------------
' IsShapeOnPolyLineCurve (Function), source lines 682-709
' -----------------------------------------------------------------------------
Private Function IsShapeOnPolyLineCurve(crv As VGCore.curve, x As Double, y As Double, shape As VGCore.shape) As Boolean
    With shape
        If .Type = cdrCurveShape Then
            If .curve.closed Then
                If IsOnPolyLineCurve(crv, x, y) Then
                    IsShapeOnPolyLineCurve = True
                Else
                    IsShapeOnPolyLineCurve = IsOnPolyLineCurve(crv, .curve.Nodes.First.PositionX, .curve.Nodes.First.PositionY)
                End If

            Else
                If IsOnPolyLineCurve(crv, .curve.Nodes.First.PositionX, .curve.Nodes.First.PositionY) Then
                    IsShapeOnPolyLineCurve = True
                Else
                    If IsOnPolyLineCurve(crv, .curve.Nodes.Last.PositionX, .curve.Nodes.Last.PositionY) Then
                        IsShapeOnPolyLineCurve = True
                    Else
                        IsShapeOnPolyLineCurve = IsOnPolyLineCurve(crv, x, y)
                    End If
                End If
            End If
        ElseIf .Type = cdrGroupShape Then
            IsShapeOnPolyLineCurve = IsShapeOnPolyLineCurve(crv, x, y, .shapes(1))
        Else
            IsShapeOnPolyLineCurve = IsOnPolyLineCurve(crv, x, y)
        End If
    End With
End Function

' -----------------------------------------------------------------------------
' IsOnPolyLineCurve (Function), source lines 710-745
' -----------------------------------------------------------------------------
Private Function IsOnPolyLineCurve(crv As VGCore.curve, x As Double, y As Double) As Boolean
    Dim q As Long, q2 As Long
    Dim x1 As Double, y1 As Double
    Dim x2 As Double, y2 As Double
    Dim dy As Double
    
    With crv
        For q = 1 To .Subpaths.count
            With .Subpaths(q)
                If .closed Then
                    If x > .boundingBox.left And x < .boundingBox.right Then
                        If y > .boundingBox.bottom And y < .boundingBox.top Then
                            For q2 = 1 To .segments.count
                                GetX1Y1X2Y2 x1, y1, x2, y2, .segments(q2)
                                If y > y1 Then
                                    If y <= y2 Then
                                        With .segments(q2)
                                            If x > .StartNode.PositionX Or x > .EndNode.PositionX Then
                                                dy = (y2 - y1)
                                                If dy > 0 Then
                                                    If (y - y1) * (x2 - x1) / dy + x1 < x Then
                                                        IsOnPolyLineCurve = Not IsOnPolyLineCurve
                                                    End If
                                                End If
                                            End If
                                        End With
                                    End If
                                End If
                            Next
                        End If
                    End If
                End If
            End With
        Next
    End With
End Function

' -----------------------------------------------------------------------------
' GetX1Y1X2Y2 (Sub), source lines 746-768
' -----------------------------------------------------------------------------
Private Sub GetX1Y1X2Y2(ByRef x1 As Double, ByRef y1 As Double, ByRef x2 As Double, ByRef y2 As Double, segment As VGCore.segment)
    With segment
        If .StartNode.PositionY < .EndNode.PositionY Then
            With .StartNode
                x1 = .PositionX
                y1 = .PositionY
            End With
            With .EndNode
                x2 = .PositionX
                y2 = .PositionY
            End With
        Else
            With .EndNode
                x1 = .PositionX
                y1 = .PositionY
            End With
            With .StartNode
                x2 = .PositionX
                y2 = .PositionY
            End With
        End If
    End With
End Sub

' -----------------------------------------------------------------------------
' DeleteShapeFromShapesArray (Sub), source lines 769-775
' -----------------------------------------------------------------------------
Private Sub DeleteShapeFromShapesArray(ShapesArray() As VGCore.shape, ByVal q As Long)
    Dim q2 As Long
    For q2 = q To UBound(ShapesArray) - 1
        Set ShapesArray(q2) = ShapesArray(q2 + 1)
    Next
    ReDim Preserve ShapesArray(UBound(ShapesArray) - 1)
End Sub
