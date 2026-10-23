' =============================================================================
' Group: 03_Geometry_and_Curves
' Source module: RemoveDuplicateLines.bas
' Procedures: 28
' =============================================================================

' -----------------------------------------------------------------------------
' Remove (Sub), source lines 36-113
' -----------------------------------------------------------------------------
Public Sub Remove()
    If Processing Then Exit Sub
    UserData.Load
    LastRefreshTime = Timer
    StartTime = LastRefreshTime
    Processing = True
    ProgressFormIsShowed = False
    ThresholdAngle = 1.74532925199433E-02 * 5
    
    lastUnit = Application.ActiveDocument.Unit
    On Error GoTo er:
    Application.ActiveDocument.Unit = cdrMillimeter
    Dim RegData1 As New RegData

    Sensitivity = CDbl(RegData1.Sensitivity) / 1000
    OutlineUse = RegData1.OutlineUse

    Aborted = False
    Application.Status.BeginProgress
    Application.Optimization = True
    
    Application.ActiveDocument.PreserveSelection = True

    Dim srToRemoveDuplicates As VGCore.ShapeRange
    Set srToRemoveDuplicates = GetShapesRage(VGCore.ActiveSelectionRange)
        
    If srToRemoveDuplicates.count > 0 Then
        On Error GoTo e:
        Application.ActiveDocument.BeginCommandGroup localization.ActiveLng.Item("RemoveDuplicateLines.Command")
        Dim Curves() As VGCore.curve
        currentOperationNumber = 0
        OperationCount = 0

        Curves = GetCurves(srToRemoveDuplicates)
        Dim sr As New ShapeRange
        sr.AddRange VGCore.ActiveSelectionRange
        Application.ActiveDocument.ClearSelection
        
        
        t = Timer
'        Dim t2 As Double
'        t2 = Timer
        RemoveCoincidentCurvesSelfy Curves()
        
        ProgressForm.Percent = 1
        Application.Status.Progress = 50
        currentOperationNumber = 0
        OperationCount = (UBound(Curves()) + 1) * UBound(Curves()) \ 2
        RemoveCoincidentCurves Curves()
        
        ProgressForm.Percent = 1
        Application.Status.Progress = 100
        
        DeleteZerro Curves()
        Connecting Curves()
        closing Curves()
        sr.All.AddToSelection
        
e:
        ProgressForm.MbClosing = True
        Unload ProgressForm
        Application.ActiveDocument.EndCommandGroup
    End If

er:
    On Error Resume Next
    Application.ActiveDocument.Unit = lastUnit
    Application.ActiveDocument.PreserveSelection = False
    
    If Aborted Then Application.ActiveDocument.Undo
    Application.Optimization = False
    Application.Status.EndProgress
    Processing = False
    Application.Refresh

'    MsgBox Timer - t2

End Sub

' -----------------------------------------------------------------------------
' GetShapesRage (Function), source lines 114-131
' -----------------------------------------------------------------------------
Private Function GetShapesRage(ShapeRage As VGCore.ShapeRange) As VGCore.ShapeRange
    Dim GSR As New VGCore.ShapeRange
    Dim q As Long
    For q = 1 To ShapeRage.count
        With ShapeRage
            Select Case .shapes(q).Type
                Case cdrRectangleShape, cdrEllipseShape, cdrPolygonShape, cdrPerfectShape
                    .shapes(q).ConvertToCurves
                    GSR.Add .shapes(q)
                Case cdrShapeType.cdrCurveShape
                    GSR.Add .shapes(q)
                Case cdrShapeType.cdrGroupShape
                    GSR.AddRange GetShapesRage(.shapes(q).shapes.All)
            End Select
        End With
    Next
    Set GetShapesRage = GSR
End Function

' -----------------------------------------------------------------------------
' GetCurves (Function), source lines 133-177
' -----------------------------------------------------------------------------
Private Function GetCurves(ShapeRage As VGCore.ShapeRange) As VGCore.curve()
    Dim Curves() As VGCore.curve
    Dim txt As String
    txt = localization.ActiveLng.Item("RemoveDuplicateLines.Command") & " (1 / 3)"
    ProgressForm.StageName = txt
    
    ReDim Curves(ShapeRage.count - 1)
    ReDim OutlineWidths(ShapeRage.count - 1)
    ShapeRage.Sort " @shape1.Width + @shape1.Height > @shape2.Width + @shape2.Height"
    
    Dim q As Long
    For q = 1 To ShapeRage.count
        With ShapeRage(q)
            OutlineWidths(q - 1) = .Outline.Width / 2
            If OutlineWidths(q - 1) < 0.025 Then OutlineWidths(q - 1) = 0.025
            Set Curves(q - 1) = .curve
            Curves(q - 1).closed = False
            With .curve.segments
                OperationCount = OperationCount + .count * (.count - 1) \ 2
            End With
        End With
        
        '------Progress-Bar-------
        If ProgressFormIsShowed Then
            If Timer - LastRefreshTime > 0.1 Then
                ProgressForm.Percent = q / ShapeRage.count
                LastRefreshTime = Timer
                
                DoEvents
                If ProgressForm.Break Then
                    Aborted = True
                    Exit Function
                End If
            End If
        Else
            If Timer - StartTime > 0.5 Then
                ProgressFormIsShowed = True
                ProgressForm.Show
            End If
        End If
        '------Progress-Bar-------
        
    Next
    GetCurves = Curves()
End Function

' -----------------------------------------------------------------------------
' RemoveCoincidentCurvesSelfy (Sub), source lines 179-190
' -----------------------------------------------------------------------------
Private Sub RemoveCoincidentCurvesSelfy(Curves() As VGCore.curve)
    Dim txt As String
    txt = localization.ActiveLng.Item("RemoveDuplicateLines.Command") & " (2 / 3)"
    ProgressForm.StageName = txt
    
    Dim q As Long
    For q = 0 To UBound(Curves)
        If OutlineUse Then Sensitivity = OutlineWidths(q)
        RemoveCoincidentCurveSelfy Curves(q)
        If Aborted Then Exit Sub
    Next
End Sub

' -----------------------------------------------------------------------------
' RemoveCoincidentCurveSelfy (Sub), source lines 191-263
' -----------------------------------------------------------------------------
Private Sub RemoveCoincidentCurveSelfy(curve As VGCore.curve)

    Dim left As Double
    Dim right As Double
    Dim top As Double
    Dim bottom As Double
    Dim l As Double
    
    Dim q1 As Long, q2 As Long
    Dim RCSTS As Long
    
    q1 = 1
    With curve.segments
        Do While q1 <= .count - 1
            With .Item(q1)
                With .boundingBox
                    left = .left - Sensitivity
                    right = .right + Sensitivity
                    top = .top + Sensitivity
                    bottom = .bottom - Sensitivity
                End With
                l = .Length
            End With
            
            q2 = q1 + 1
            Do While q2 <= .count
                currentOperationNumber = currentOperationNumber + 1
                If Timer - t > 0.2 Then
                    t = Timer
                    Application.Status.Progress = CLng(currentOperationNumber / OperationCount * 100) / 2
                    If Application.Status.Aborted Then
                        Aborted = True
                        Exit Sub
                    End If
                End If
                
                '------Progress-Bar-------
                If ProgressFormIsShowed Then
                    If Timer - LastRefreshTime > 0.2 Then
                        ProgressForm.Percent = currentOperationNumber / OperationCount
                        LastRefreshTime = Timer
                        
                        DoEvents
                        If ProgressForm.Break Then
                            Aborted = True
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
                
                
                If BoundingBoxCrossTest(.Item(q2).boundingBox, left, right, top, bottom) Then
                    If l >= .Item(q2).Length Then
                        q2 = q2 + RemoveCoincidentSegmentToSegment(.Item(q1), .Item(q2))
                    Else
                        RCSTS = RemoveCoincidentSegmentToSegment(.Item(q2), .Item(q1))
                        q1 = q1 + RCSTS
                        If RCSTS < 0 Then GoTo ExitLoop
                    End If
                End If
                q2 = q2 + 1
            Loop
ExitLoop:
            q1 = q1 + 1
        Loop
    End With
End Sub

' -----------------------------------------------------------------------------
' BoundingBoxCrossTest (Function), source lines 264-276
' -----------------------------------------------------------------------------
Private Function BoundingBoxCrossTest(ByRef bBox As VGCore.rect, ByRef left As Double, ByRef right As Double, ByRef top As Double, ByRef bottom As Double) As Boolean
    With bBox
        If left <= .right Then
            If right >= .left Then
                If top >= .bottom Then
                    If bottom <= .top Then
                        BoundingBoxCrossTest = True
                    End If
                End If
            End If
        End If
    End With
End Function

' -----------------------------------------------------------------------------
' RemoveCoincidentCurves (Sub), source lines 279-352
' -----------------------------------------------------------------------------
Private Sub RemoveCoincidentCurves(Curves() As VGCore.curve)
    Dim q1 As Long, q2 As Long, ub As Long
    ub = UBound(Curves)
    
    Dim left As Double
    Dim right As Double
    Dim top As Double
    Dim bottom As Double

    Dim txt As String
    txt = localization.ActiveLng.Item("RemoveDuplicateLines.Command") & " (3 / 3)"
    ProgressForm.StageName = txt

    For q1 = 0 To ub - 1
        With Curves(q1)
            If .segments.count <> 0 Then
                With .boundingBox
                    left = .left - Sensitivity
                    right = .right + Sensitivity
                    top = .top + Sensitivity
                    bottom = .bottom - Sensitivity
                End With
                
                For q2 = q1 + 1 To ub
                    With Curves(q2)
                        If .segments.count <> 0 Then
                            currentOperationNumber = currentOperationNumber + 1
                            If Timer - t > 0.2 Then
                                t = Timer
                                Application.Status.Progress = CLng(currentOperationNumber / OperationCount * 100) \ 2 + 50

                                If Application.Status.Aborted Then
                                    Aborted = True
                                    Exit Sub
                                End If
                            End If
                            
                            '------Progress-Bar-------
                            If ProgressFormIsShowed Then
                                If Timer - LastRefreshTime > 0.2 Then
                                    ProgressForm.Percent = currentOperationNumber / OperationCount
                                    LastRefreshTime = Timer
                                     
                                    DoEvents
                                    If ProgressForm.Break Then
                                        Aborted = True
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
                            
                            If BoundingBoxCrossTest(.boundingBox, left, right, top, bottom) Then
                                If OutlineUse Then
                                    If OutlineWidths(q1) > OutlineWidths(q2) Then
                                        Sensitivity = OutlineWidths(q1)
                                    Else
                                        Sensitivity = OutlineWidths(q2)
                                    End If
                                End If
                                RemoveCoincidentCurveToCurve Curves(q1).segments, Curves(q2)
                            End If
                        End If
                    End With
                Next
            End If
        End With
    Next
End Sub

' -----------------------------------------------------------------------------
' RemoveCoincidentCurveToCurve (Sub), source lines 353-379
' -----------------------------------------------------------------------------
Private Sub RemoveCoincidentCurveToCurve(TrimmingSegments As VGCore.segments, BreakCurve As VGCore.curve)

    Dim BreakSegments As VGCore.segments
    Set BreakSegments = BreakCurve.segments

    With BreakCurve.boundingBox
        Dim left As Double
        Dim right As Double
        Dim top As Double
        Dim bottom As Double
        left = .left - Sensitivity
        right = .right + Sensitivity
        top = .top + Sensitivity
        bottom = .bottom - Sensitivity
    End With
    
    Dim q As Long
    q = 1
    With TrimmingSegments
        Do While q <= .count
            If BoundingBoxCrossTest(TrimmingSegments(q).boundingBox, left, right, top, bottom) Then
                RemoveCoincidentSegmentToCurve TrimmingSegments(q), BreakSegments
            End If
            q = q + 1
        Loop
    End With
End Sub

' -----------------------------------------------------------------------------
' RemoveCoincidentSegmentToCurve (Sub), source lines 381-445
' -----------------------------------------------------------------------------
Private Sub RemoveCoincidentSegmentToCurve(TrimmingSegment As VGCore.segment, BreakSegments As VGCore.segments)

    With TrimmingSegment.boundingBox
        Dim left As Double
        Dim right As Double
        Dim top As Double
        Dim bottom As Double
        left = .left - Sensitivity
        right = .right + Sensitivity
        top = .top + Sensitivity
        bottom = .bottom - Sensitivity
    End With
     
    Dim RCSTS As Long
    Dim q As Long
    
    If TrimmingSegment.Type = cdrCurveSegment Then
        Dim PolylineCreated As Boolean
        Dim TrimmingPolyLine As VGCore.curve

        With BreakSegments
            q = 1
            Do While q <= .count
                If BoundingBoxCrossTest(BreakSegments(q).boundingBox, left, right, top, bottom) Then
                    If BreakSegments(q).Type = cdrCurveSegment Then
                        'Bezier-Bezier
                        RCSTS = RemoveCoincidentBezierSegmentToBezierSegment(TrimmingSegment, BreakSegments(q))
                        q = q + RCSTS
                        If RCSTS > 0 Then Exit Sub
                    Else
                        'Bezier-Line
                        If Not PolylineCreated Then
                            Set TrimmingPolyLine = CreatePolylineCurve(TrimmingSegment)
                            PolylineCreated = True
                        End If
                        RCSTS = RemoveCoincidentBezierSegmentToLineSegment(TrimmingPolyLine, BreakSegments(q))
                        q = q + RCSTS
                        If RCSTS > 0 Then Exit Sub
                    End If
                End If
                q = q + 1
            Loop
        End With
    Else
        With BreakSegments
            q = 1
            Do While q <= .count
                If BoundingBoxCrossTest(BreakSegments(q).boundingBox, left, right, top, bottom) Then
                    If BreakSegments(q).Type = cdrCurveSegment Then
                        'Line-Bezier
                        RCSTS = RemoveCoincidentLineSegmentToBezierSegment(TrimmingSegment, BreakSegments(q))
                        q = q + RCSTS
                        If RCSTS > 0 Then Exit Sub
                    Else
                        'Line-Line
                        RCSTS = RemoveCoincidentLineSegmentToLineSegment(TrimmingSegment, BreakSegments(q))
                        q = q + RCSTS
                        If RCSTS > 0 Then Exit Sub
                    End If
                End If
                q = q + 1
            Loop
        End With
    End If
End Sub

' -----------------------------------------------------------------------------
' RemoveCoincidentLineSegmentToLineSegment (Function), source lines 447-544
' -----------------------------------------------------------------------------
Private Function RemoveCoincidentLineSegmentToLineSegment(TrimmingSegment As VGCore.segment, BreakSegment As VGCore.segment) As Long
    'Line-Line
    'x--------->  BreakSegment
    'x=========>  TrimmingSegment

    If (DeltaAnglBetweenSegments(BreakSegment, TrimmingSegment)) < ThresholdAngle Then
        Dim PT As ParametrT2X
        Dim dist1 As Distance2X
        Dim dist2 As Distance2X
        
        PT = ParametrT(BreakSegment, TrimmingSegment)
        dist1 = Distance(BreakSegment, TrimmingSegment)
        dist2 = Distance(TrimmingSegment, BreakSegment)
        Dim delNode As VGCore.Node
        Dim t As Double
         
        With PT
            If .t2 > .t1 Then
                '>>>>>>
                If dist1.Distance1 < Sensitivity Or dist2.Distance1 < Sensitivity Then
                    If dist1.Distance2 < Sensitivity Or dist2.Distance2 < Sensitivity Then
                        Select Case ((PT.t1 > 0 And PT.t1 < 1) * 2 + (PT.t2 > 0 And PT.t2 < 1))
                            Case 0 'false false
                                'x--------->  x====>, x====>   x------->
                                If PT.t1 <= 0 And PT.t2 >= 1 Then
                                    'x====x====>======>, xx===>>
                                    Set delNode = BreakSegment.EndNode
                                    BreakSegment.EndNode.BreakApart
                                    delNode.Delete
                                    RemoveCoincidentLineSegmentToLineSegment = -1
                                End If
                            Case -1 'false true
                                'x==x=>------->
                                Set delNode = BreakSegment.StartNode
                                BreakSegment.BreakApartAt PT.t2
                                delNode.Next.Delete
                                'RemoveCoincidentLineSegmentToLineSegment = 0
                            Case -2 'true false
                                'x------x==>==>
                                Set delNode = BreakSegment.EndNode
                                BreakSegment.BreakApartAt PT.t1
                                delNode.Next.Delete
                                'RemoveCoincidentLineSegmentToLineSegment = 0
                            Case -3 'true true
                                'x--x====>->
                                t = (PT.t2 - PT.t1) / (1 - PT.t1)
                                If t > 0 And t < 1 Then
                                    Set delNode = BreakSegment.BreakApartAt(PT.t1)
                                    delNode.NextSegment.BreakApartAt t
                                    delNode.Delete
                                    RemoveCoincidentLineSegmentToLineSegment = 1
                                End If
                        End Select
                    End If
                End If
            Else
                '>>><<<
                If dist1.Distance1 < Sensitivity Or dist2.Distance2 < Sensitivity Then
                    If dist1.Distance2 < Sensitivity Or dist2.Distance1 < Sensitivity Then
                        Select Case (PT.t2 > 0 And PT.t2 < 1) * 2 + (PT.t1 > 0 And PT.t1 < 1)
                            Case 0 'false false
                                'x--------->  <====x, x====>   <-------x
                                If PT.t2 <= 0 And PT.t1 >= 1 Then
                                    '<====x====>======x, <x===>x
                                    Set delNode = BreakSegment.EndNode
                                    BreakSegment.EndNode.BreakApart
                                    delNode.Delete
                                    RemoveCoincidentLineSegmentToLineSegment = -1
                                End If
                            Case -1 'false true
                                '<==x===x----->
                                Set delNode = BreakSegment.StartNode
                                BreakSegment.BreakApartAt PT.t1
                                BreakSegment.EndNode.Delete
'                                delNode.Next.Delete
                                'RemoveCoincidentLineSegmentToLineSegment = 0
                            Case -2 'true false
                                'x------<==>===x
                                Set delNode = BreakSegment.EndNode
                                BreakSegment.BreakApartAt PT.t2
                                delNode.Next.Delete
                                'RemoveCoincidentLineSegmentToLineSegment = 0
                            Case -3 'true true
                                'x--<====x->
                                t = (PT.t1 - PT.t2) / (1 - PT.t2)
                                If t > 0 And t < 1 Then
                                    Set delNode = BreakSegment.BreakApartAt(PT.t2)
                                    delNode.NextSegment.BreakApartAt t
                                    delNode.Delete
                                    RemoveCoincidentLineSegmentToLineSegment = 1
                                End If
                        End Select
                    End If
                End If
            End If
        End With
    End If
End Function

' -----------------------------------------------------------------------------
' DeltaAnglBetweenSegments (Function), source lines 545-554
' -----------------------------------------------------------------------------
Private Function DeltaAnglBetweenSegments(Segment1 As VGCore.segment, Segment2 As VGCore.segment) As Double
    On Error GoTo Err
    Static d As Double
    d = Abs(Angel(Segment1) - Angel(Segment2))
    If d > 1.5707963267949 Then
        d = 3.14159265358979 - d
    End If
    DeltaAnglBetweenSegments = d
Err:
End Function

' -----------------------------------------------------------------------------
' Angel (Function), source lines 555-579
' -----------------------------------------------------------------------------
Private Function Angel(segment As VGCore.segment) As Double
    Static x1 As Double
    Static y1 As Double
    Static x2 As Double
    Static y2 As Double
    
    With segment
        With .StartNode
            x1 = .PositionX
            y1 = .PositionY
        End With
        With .EndNode
            x2 = .PositionX
            y2 = .PositionY
        End With
    End With

    If x2 <> x1 Then
        Angel = Atn((y2 - y1) / (x2 - x1))
    ElseIf y2 <> y1 Then
        Angel = 1.5707963267949
    Else
        Err.Raise vbObjectError, , ""
    End If
End Function

' -----------------------------------------------------------------------------
' ParametrT (Function), source lines 580-609
' -----------------------------------------------------------------------------
Private Function ParametrT(segment As VGCore.segment, SegmentForPoints As VGCore.segment) As ParametrT2X
    Static dx As Double
    Static dy As Double
    Static dPx1 As Double
    Static dPy1 As Double
    Static dPx2 As Double
    Static dPy2 As Double
    
    With segment
        dx = .EndNode.PositionX - .StartNode.PositionX
        dy = .EndNode.PositionY - .StartNode.PositionY
    End With

    With SegmentForPoints
        dPx1 = .StartNode.PositionX - segment.StartNode.PositionX
        dPy1 = .StartNode.PositionY - segment.StartNode.PositionY
        dPx2 = .EndNode.PositionX - segment.StartNode.PositionX
        dPy2 = .EndNode.PositionY - segment.StartNode.PositionY
    End With

    With ParametrT
        If Abs(dx) > Abs(dy) Then
            .t1 = dPx1 / dx
            .t2 = dPx2 / dx
        ElseIf dy <> 0 Then
            .t1 = dPy1 / dy
            .t2 = dPy2 / dy
        End If
    End With
End Function

' -----------------------------------------------------------------------------
' Distance (Function), source lines 611-661
' -----------------------------------------------------------------------------
Private Function Distance(segment As VGCore.segment, SegmentForPoints As VGCore.segment) As Distance2X
    Static dx As Double
    Static dy As Double
    Static x1 As Double
    Static y1 As Double
    Static x2 As Double
    Static y2 As Double
    Static l As Double
    Static px1 As Double
    Static py1 As Double
    Static px2 As Double
    Static py2 As Double
    Static tmp As Double
    
    With SegmentForPoints
        With .StartNode
            px1 = .PositionX
            py1 = .PositionY
        End With
        With .EndNode
            px2 = .PositionX
            py2 = .PositionY
        End With
    End With

    With segment
        With .StartNode
            x1 = .PositionX
            y1 = .PositionY
        End With
        With .EndNode
            x2 = .PositionX
            y2 = .PositionY
        End With
    End With
    dx = x2 - x1
    dy = y2 - y1
    l = Sqr(dy * dy + dx * dx)
    If l > 0 Then
        With Distance
            tmp = x2 * y1 - y2 * x1
            .Distance1 = Abs(dy * px1 - dx * py1 + tmp) / l
            .Distance2 = Abs(dy * px2 - dx * py2 + tmp) / l
        End With
    Else
        With Distance
            .Distance1 = Sqr((py1 - y1) * (py1 - y1) + (px1 - x1) * (px1 - x1))
            .Distance2 = Sqr((py2 - y1) * (py2 - y1) + (px2 - x1) * (px2 - x1))
        End With
    End If
End Function

' -----------------------------------------------------------------------------
' RemoveCoincidentSegmentToSegment (Function), source lines 662-684
' -----------------------------------------------------------------------------
Private Function RemoveCoincidentSegmentToSegment(TrimmingSegment As VGCore.segment, BreakSegment As VGCore.segment) As Long
    If TrimmingSegment.Type = cdrLineSegment Then
        If BreakSegment.Type = cdrLineSegment Then
            'Line-Line
            RemoveCoincidentSegmentToSegment = RemoveCoincidentLineSegmentToLineSegment(TrimmingSegment, BreakSegment)
        ElseIf BreakSegment.Type = cdrCurveSegment Then
            'Line-Bezier
            RemoveCoincidentSegmentToSegment = RemoveCoincidentLineSegmentToBezierSegment(TrimmingSegment, BreakSegment)
        End If
    ElseIf TrimmingSegment.Type = cdrCurveSegment Then
        If BreakSegment.Type = cdrLineSegment Then
            'Bezier-Line
            Dim TrimmingPolyLine As VGCore.curve
            Set TrimmingPolyLine = CreatePolylineCurve(TrimmingSegment)
            RemoveCoincidentSegmentToSegment = RemoveCoincidentBezierSegmentToLineSegment(TrimmingPolyLine, BreakSegment)
        ElseIf BreakSegment.Type = cdrCurveSegment Then
            'Bezier-Bezier
            If Not RemoveCoincidentBezierSegmentToBezierSegmentQuickTest(TrimmingSegment, BreakSegment) Then
                RemoveCoincidentSegmentToSegment = RemoveCoincidentBezierSegmentToBezierSegment(TrimmingSegment, BreakSegment)
            End If
        End If
    End If
End Function

' -----------------------------------------------------------------------------
' DeleteZerro (Sub), source lines 685-698
' -----------------------------------------------------------------------------
Private Sub DeleteZerro(Curves() As VGCore.curve)
    Dim q As Long, q2 As Long
    For q = 0 To UBound(Curves)
        With Curves(q)
            If .segments.count > 0 Then
                For q2 = .Subpaths.count To 1 Step -1
                    With .Subpaths(q2)
                        If .Length = 0 Then .Delete
                    End With
                Next
            End If
        End With
    Next
End Sub

' -----------------------------------------------------------------------------
' Connecting (Sub), source lines 699-729
' -----------------------------------------------------------------------------
Private Sub Connecting(Curves() As VGCore.curve)
    Dim q As Long, q1 As Long, q2 As Long
    For q = 0 To UBound(Curves)
        With Curves(q)
            If .segments.count > 0 Then
                q1 = 1
                Do While q1 <= .Subpaths.count
                    q2 = 1
                    Do While q2 <= .Subpaths.count
                        If q1 <> q2 Then
                            If .Subpaths(q1).EndNode.PositionX = .Subpaths(q2).StartNode.PositionX Then
                                If .Subpaths(q1).EndNode.PositionY = .Subpaths(q2).StartNode.PositionY Then
                                    .Subpaths(q1).EndNode.ConnectWith .Subpaths(q2).StartNode
                                    If q1 > q2 Then
                                        q1 = q1 - 2
                                    Else
                                        q1 = q1 - 1
                                    End If
                                    GoTo ExitLoop
                                End If
                            End If
                        End If
                        q2 = q2 + 1
                    Loop
ExitLoop:
                    q1 = q1 + 1
                Loop
            End If
        End With
    Next
End Sub

' -----------------------------------------------------------------------------
' closing (Sub), source lines 730-747
' -----------------------------------------------------------------------------
Private Sub closing(Curves() As VGCore.curve)
    Dim q As Long, q2 As Long
    For q = 0 To UBound(Curves)
        With Curves(q)
            If .segments.count > 0 Then
                For q2 = 1 To .Subpaths.count
                    With .Subpaths(q2)
                        If .StartNode.PositionX = .EndNode.PositionX Then
                            If .StartNode.PositionY = .EndNode.PositionY Then
                                .closed = True
                            End If
                        End If
                    End With
                Next
            End If
        End With
    Next
End Sub

' -----------------------------------------------------------------------------
' CreatePolylineCurve (Function), source lines 757-770
' -----------------------------------------------------------------------------
Private Function CreatePolylineCurve(BezierSegment As VGCore.segment, Optional ByRef minLength As Double) As VGCore.curve
    With BezierSegment
        minLength = .Length
        Dim CountPolyLines As Long
        CountPolyLines = minLength * 0.2
        If CountPolyLines < 50 Then
            CountPolyLines = 50
        ElseIf CountPolyLines > 500 Then
            CountPolyLines = 500
        End If
        Set CreatePolylineCurve = .GetPolyline(CountPolyLines)
        minLength = minLength / CountPolyLines / 100
    End With
End Function

' -----------------------------------------------------------------------------
' CreatePolylineCurve2 (Function), source lines 771-773
' -----------------------------------------------------------------------------
Private Function CreatePolylineCurve2(BezierSegment As VGCore.segment, CountPolyLines As Long) As VGCore.curve
    Set CreatePolylineCurve2 = BezierSegment.GetPolyline(CountPolyLines)
End Function

' -----------------------------------------------------------------------------
' RemoveCoincidentBezierSegmentToBezierSegmentQuickTest (Function), source lines 774-798
' -----------------------------------------------------------------------------
Private Function RemoveCoincidentBezierSegmentToBezierSegmentQuickTest(TrimmingSegment As VGCore.segment, BreakSegment As VGCore.segment) As Boolean
    With TrimmingSegment
        If .StartNode.PositionX = BreakSegment.StartNode.PositionX Then
            If .StartNode.PositionY = BreakSegment.StartNode.PositionY Then
                If .EndNode.PositionX = BreakSegment.EndNode.PositionX Then
                    If .EndNode.PositionY = BreakSegment.EndNode.PositionY Then
                        If .StartingControlPointX = BreakSegment.StartingControlPointX Then
                            If .StartingControlPointY = BreakSegment.StartingControlPointY Then
                                If .EndingControlPointX = BreakSegment.EndingControlPointX Then
                                    If .EndingControlPointY = BreakSegment.EndingControlPointY Then
                                        RemoveCoincidentBezierSegmentToBezierSegmentQuickTest = True
                                        With BreakSegment.EndNode
                                            .BreakApart
                                            .Delete
                                        End With
                                    End If
                                End If
                            End If
                        End If
                    End If
                End If
            End If
        End If
    End With
End Function

' -----------------------------------------------------------------------------
' RemoveCoincidentBezierSegmentToBezierSegment (Function), source lines 801-873
' -----------------------------------------------------------------------------
Private Function RemoveCoincidentBezierSegmentToBezierSegment(TrimmingSegment As VGCore.segment, BreakSegment As VGCore.segment) As Long
    'Bezier-Bezier
    Dim OldThresholdAngle As Double
    Dim OldSensitivity As Double
    OldThresholdAngle = ThresholdAngle
    OldSensitivity = Sensitivity
    
    Sensitivity = Sensitivity * 1.15
    ThresholdAngle = ThresholdAngle * 5
    
    Dim nc As Long
    Dim mSegments As VGCore.segments
    Set mSegments = BreakSegment.Parent.segments
    nc = mSegments.count
    
    Dim minLength As Double
    Dim BreakPolyLine As VGCore.curve
    Dim TrimmingPolyLine As VGCore.curve
    
    Set BreakPolyLine = CreatePolylineCurve(BreakSegment, minLength)
    Dim c As Double
    If minLength > 0 Then
        c = TrimmingSegment.Length / (minLength * 100)
        If c > 500 Then c = 500
    Else
        c = 500
    End If
    Set TrimmingPolyLine = CreatePolylineCurve2(TrimmingSegment, CLng(c))
    
    Dim rs() As RemovableSegment

    RemoveCoincidentCurveToCurve TrimmingPolyLine.segments, BreakPolyLine
    DeleteLittleSubPathsFromPolyLine BreakPolyLine.Subpaths, minLength
    With BreakSegment
        rs() = GetInvertSegments(BreakPolyLine.Subpaths, .StartNode.PositionX, .StartNode.PositionY, .EndNode.PositionX, .EndNode.PositionY)
    End With
    
    Dim q As Long
    Dim t As Double
    Dim NewNode As VGCore.Node
        
    If Not Not rs() Then
        For q = 0 To UBound(rs)
            With rs(q)
                'Start Cut Line
                t = GetParametrT(.startX, .startY, BreakSegment)
                If t > 0 And t < 1 Then
                    Set NewNode = BreakSegment.AddNodeAt(t) '(this is node from next segment)
                    Set BreakSegment = NewNode.NextSegment
                End If
                
                'End Cut Line
                t = GetParametrT(.EndX, .EndY, BreakSegment)
                If t > 0 And t < 1 Then
                    Set NewNode = BreakSegment.BreakApartAt(t) '(this is node from next segment)
                    BreakSegment.EndNode.Delete
                    SetNextSegment NewNode, BreakSegment
                Else
                    With BreakSegment.EndNode
                        .BreakApart
                        .Delete
                        Exit For
                    End With
                End If
            End With
        Next
    End If
    
    RemoveCoincidentBezierSegmentToBezierSegment = mSegments.count - nc
    
    ThresholdAngle = OldThresholdAngle
    Sensitivity = OldSensitivity
End Function

' -----------------------------------------------------------------------------
' RemoveCoincidentBezierSegmentToLineSegment (Function), source lines 875-932
' -----------------------------------------------------------------------------
Private Function RemoveCoincidentBezierSegmentToLineSegment(TrimmingPolyLine As VGCore.curve, BreakSegment As VGCore.segment) As Long
    'Bezier-Line
    Dim OldThresholdAngle As Double
    OldThresholdAngle = ThresholdAngle
    ThresholdAngle = ThresholdAngle * 5
    
    Dim nc As Long
    Dim mSegments As VGCore.segments
    Set mSegments = BreakSegment.Parent.segments
    nc = mSegments.count
    
    Dim BreakPolyLine  As New VGCore.curve
    Dim sp As VGCore.SubPath
    Dim rs() As RemovableSegment
    With BreakSegment
        Set sp = BreakPolyLine.CreateSubPath(.StartNode.PositionX, .StartNode.PositionY)
        sp.AppendLineSegment .EndNode.PositionX, .EndNode.PositionY
    End With
    RemoveCoincidentCurveToCurve TrimmingPolyLine.segments, BreakPolyLine
    With BreakSegment
        rs() = GetInvertSegments(BreakPolyLine.Subpaths, .StartNode.PositionX, .StartNode.PositionY, .EndNode.PositionX, .EndNode.PositionY)
    End With


    Dim q As Long
    Dim t As Double
    Dim NewNode As VGCore.Node
        
    If Not Not rs() Then
        For q = 0 To UBound(rs)
            With rs(q)
                'Start Cut Line
                t = GetParametrT(.startX, .startY, BreakSegment)
                If t > 0 And t < 1 Then
                    Set NewNode = BreakSegment.AddNodeAt(t) '(this is node from next segment)
                    Set BreakSegment = NewNode.NextSegment
                End If
                
                'End Cut Line
                t = GetParametrT(.EndX, .EndY, BreakSegment)
                If t > 0 And t < 1 Then
                    Set NewNode = BreakSegment.BreakApartAt(t) '(this is node from next segment)
                    BreakSegment.EndNode.Delete
                    SetNextSegment NewNode, BreakSegment
                Else
                    With BreakSegment.EndNode
                        .BreakApart
                        .Delete
                        Exit For
                    End With
                End If
            End With
        Next
    End If

    RemoveCoincidentBezierSegmentToLineSegment = mSegments.count - nc
    ThresholdAngle = OldThresholdAngle
End Function

' -----------------------------------------------------------------------------
' RemoveCoincidentLineSegmentToBezierSegment (Function), source lines 934-991
' -----------------------------------------------------------------------------
Private Function RemoveCoincidentLineSegmentToBezierSegment(TrimmingSegment As VGCore.segment, BreakSegment As VGCore.segment) As Long
    'Line-Bezier
    Dim OldThresholdAngle As Double
    OldThresholdAngle = ThresholdAngle
    ThresholdAngle = ThresholdAngle * 5
    
    Dim nc As Long
    Dim mSegments As VGCore.segments
    Set mSegments = BreakSegment.Parent.segments
    nc = mSegments.count
    
    Dim minLength As Double
    Dim BreakPolyLine As VGCore.curve
    Set BreakPolyLine = CreatePolylineCurve(BreakSegment, minLength)

    Dim rs() As RemovableSegment
    With BreakPolyLine
        RemoveCoincidentSegmentToCurve TrimmingSegment, .segments
        DeleteLittleSubPathsFromPolyLine .Subpaths, minLength
    End With
    With BreakSegment
        rs() = GetInvertSegments(BreakPolyLine.Subpaths, .StartNode.PositionX, .StartNode.PositionY, .EndNode.PositionX, .EndNode.PositionY)
    End With
    
    Dim q As Long
    Dim t As Double
    Dim NewNode As VGCore.Node
    
    If Not Not rs() Then
        For q = 0 To UBound(rs)
            With rs(q)
                'Start Cut Line
                t = GetParametrT(.startX, .startY, BreakSegment)
                If t > 0 And t < 1 Then
                    Set NewNode = BreakSegment.AddNodeAt(t) '(this is node from next segment)
                    Set BreakSegment = NewNode.NextSegment
                End If
                
                'End Cut Line
                t = GetParametrT(.EndX, .EndY, BreakSegment)
                If t > 0 And t < 1 Then
                    Set NewNode = BreakSegment.BreakApartAt(t) '(this is node from next segment)
                    BreakSegment.EndNode.Delete
                    SetNextSegment NewNode, BreakSegment
                Else
                    With BreakSegment.EndNode
                        .BreakApart
                        .Delete
                        Exit For
                    End With
                End If
            End With
        Next
    End If

    RemoveCoincidentLineSegmentToBezierSegment = mSegments.count - nc
    ThresholdAngle = OldThresholdAngle
End Function

' -----------------------------------------------------------------------------
' SetNextSegment (Sub), source lines 992-996
' -----------------------------------------------------------------------------
Private Sub SetNextSegment(Node As VGCore.Node, ByRef BreakSegment As VGCore.segment)
    On Error GoTo 10
    Set BreakSegment = Node.NextSegment
10
End Sub

' -----------------------------------------------------------------------------
' GetInvertSegments (Function), source lines 997-1047
' -----------------------------------------------------------------------------
Private Function GetInvertSegments(Subpaths As VGCore.Subpaths, sx As Double, sy As Double, ex As Double, ey As Double) As RemovableSegment()
    Dim n As Long
    n = Subpaths.count
    Dim rs() As RemovableSegment
    ReDim rs(n)
    
    With rs(0)
        .startX = sx
        .startY = sy
    End With
    
    With rs(n)
        .EndX = ex
        .EndY = ey
    End With
    
    Dim q As Long
    For q = 1 To n
        With Subpaths(q)
            rs(q - 1).EndX = .StartNode.PositionX
            rs(q - 1).EndY = .StartNode.PositionY
            rs(q).startX = .EndNode.PositionX
            rs(q).startY = .EndNode.PositionY
        End With
    Next
    

    If rs(0).startX = rs(0).EndX And rs(0).startY = rs(0).EndY Then
        If UBound(rs) > 0 Then
            For q = 0 To UBound(rs) - 1
                rs(q) = rs(q + 1)
            Next
            ReDim Preserve rs(UBound(rs) - 1)
        Else
            Erase rs()
        End If
    End If

    
    If Not Not rs() Then
        If rs(UBound(rs)).startX = rs(UBound(rs)).EndX And rs(UBound(rs)).startY = rs(UBound(rs)).EndY Then
            If UBound(rs) > 0 Then
                ReDim Preserve rs(UBound(rs) - 1)
            Else
                Erase rs()
            End If
        End If
    End If
    
    GetInvertSegments = rs
End Function

' -----------------------------------------------------------------------------
' GetParametrT (Function), source lines 1048-1053
' -----------------------------------------------------------------------------
Private Function GetParametrT(x As Double, y As Double, segment As VGCore.segment) As Double
    Dim t As Double
    segment.FindClosestPoint x, y, t
    t = Round(t, 10)
    GetParametrT = t
End Function

' -----------------------------------------------------------------------------
' DeleteLittleSubPathsFromPolyLine (Sub), source lines 1054-1061
' -----------------------------------------------------------------------------
Private Sub DeleteLittleSubPathsFromPolyLine(PolyLineSubPaths As VGCore.Subpaths, minLength As Double)
    Dim q As Long
    For q = PolyLineSubPaths.count To 1 Step -1
        With PolyLineSubPaths(q)
            If .Length < minLength Then .Delete
        End With
    Next
End Sub
