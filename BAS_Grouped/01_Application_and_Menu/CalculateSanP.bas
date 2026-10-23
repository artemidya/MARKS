' =============================================================================
' Group: 01_Application_and_Menu
' Source module: CalculateSanP.bas
' Procedures: 14
' =============================================================================

' -----------------------------------------------------------------------------
' Start (Sub), source lines 30-72
' -----------------------------------------------------------------------------
Public Sub Start()
    On Error GoTo 10
    lastUnit = Application.ActiveDocument.Unit
    ProcessingStr = ".."
    NeedBreak = False
    LengthOfCurves = 0
    AreaOfCurves = 0
    RectAreaOfCurves = 0
    t = Timer
    CountErr = 0
    RectOnScreenCurvesCount = 0
    CountSubPaths = 0
    CountNodes = 0
    
    CurQ_ShowRectOnScreenCurves = 1
    ReDim RectOnScreenCurves(0)
    ReDim ErrOnScreenCurves(0)
    ReDim ErrOnScreenText(0)
      
    SelectionRectArea = Application.ActiveSelection.SizeHeight * Application.ActiveSelection.SizeWidth

    
    
    With Application.ActiveSelection.boundingBox
        SelectionRectOnScreenCurve.SetRectangle .left, .top, .right, .bottom
    End With
    SelectionRectOnScreenCurve.SetBrush RGB(230, 230, 255)
    
    
    CalculateShapes Application.ActiveSelectionRange.shapes
10
    ProcessingStr = ""
    If NeedBreak = True Then
        ClearAllOnScreen
        Exit Sub
    End If
    If PerimetrFormVisible = True Then PerimetrForm.UserFormRefreshData
    If SquareFormVisible = True Then
        SquareForm.UserFormRefreshData
        If SquareForm.ComboBox1.ListIndex = 1 Then ShowRectOnScreenCurves Else HideRectOnScreenCurves
        If SquareForm.ComboBox1.ListIndex = 2 Then SelectionRectOnScreenCurve.Show Else SelectionRectOnScreenCurve.Hide
    End If
End Sub

' -----------------------------------------------------------------------------
' Break (Sub), source lines 73-76
' -----------------------------------------------------------------------------
Public Sub Break()
    NeedBreak = True
    ClearAllOnScreen
End Sub

' -----------------------------------------------------------------------------
' ClearAllOnScreen (Sub), source lines 77-83
' -----------------------------------------------------------------------------
Private Sub ClearAllOnScreen()
    Erase RectOnScreenCurves
    Erase ErrOnScreenCurves
    Erase ErrOnScreenText
    RectOnScreenCurvesCount = 0
    CurQ_ShowRectOnScreenCurves = 1
End Sub

' -----------------------------------------------------------------------------
' ErrShow (Sub), source lines 84-93
' -----------------------------------------------------------------------------
Public Sub ErrShow()
    On Error GoTo 10
    Dim q As Long
    For q = 1 To UBound(ErrOnScreenCurves)
        ErrOnScreenCurves(q).Show
        ErrOnScreenText(q).Show
    Next
    Exit Sub
10:
End Sub

' -----------------------------------------------------------------------------
' ErrHide (Sub), source lines 95-104
' -----------------------------------------------------------------------------
Public Sub ErrHide()
    On Error GoTo 10
    Dim q As Long
    For q = 1 To UBound(ErrOnScreenCurves)
        ErrOnScreenCurves(q).Hide
        ErrOnScreenText(q).Hide
    Next
    Exit Sub
10:
End Sub

' -----------------------------------------------------------------------------
' ShowRectOnScreenCurves (Sub), source lines 107-113
' -----------------------------------------------------------------------------
Public Sub ShowRectOnScreenCurves()
    On Error GoTo 10
    For CurQ_ShowRectOnScreenCurves = CurQ_ShowRectOnScreenCurves To RectOnScreenCurvesCount
        RectOnScreenCurves(CurQ_ShowRectOnScreenCurves).Show
    Next
10
End Sub

' -----------------------------------------------------------------------------
' HideRectOnScreenCurves (Sub), source lines 114-122
' -----------------------------------------------------------------------------
Public Sub HideRectOnScreenCurves()
    On Error GoTo 10
    Dim q As Long
    For q = 1 To CurQ_ShowRectOnScreenCurves - 1
        RectOnScreenCurves(q).Hide
    Next
    CurQ_ShowRectOnScreenCurves = 1
10
End Sub

' -----------------------------------------------------------------------------
' CalculateShapes (Sub), source lines 124-157
' -----------------------------------------------------------------------------
Private Sub CalculateShapes(ByRef shapes As VGCore.shapes)
    Dim q1 As Long
    For q1 = 1 To shapes.count
        If shapes(q1).Type = cdrCurveShape Then
            LengthOfCurves = LengthOfCurves + Length(shapes(q1).curve)
            AreaOfCurves = AreaOfCurves + Area(shapes(q1).curve)
            RectAreaOfCurves = RectAreaOfCurves + RectArea(shapes(q1).curve)
            CountSubPaths = CountSubPaths + shapes(q1).curve.Subpaths.count
            CountNodes = CountNodes + shapes(q1).curve.Nodes.count
            
        ElseIf shapes(q1).Type = cdrGroupShape _
        Or shapes(q1).Type = cdrContourGroupShape _
        Or shapes(q1).Type = cdrArtisticMediaGroupShape Then
             CalculateShapes shapes(q1).shapes
        ElseIf shapes(q1).Type = cdrBitmapShape Then
            CountErrPlus1 shapes(q1)
        Else
            '.......................
            If TryDisplayCurve(shapes(q1)) = False Then
                CountErrPlus1 shapes(q1)
            End If
        End If
        If Timer - t > 0.02 Then
            t = Timer
            If PerimetrFormVisible = True Then PerimetrForm.UserFormRefreshData
            If SquareFormVisible = True Then
                SquareForm.UserFormRefreshData
                If SquareForm.ComboBox1.ListIndex = 1 Then ShowRectOnScreenCurves Else HideRectOnScreenCurves
            End If
            DoEvents
        End If
        If NeedBreak = True Then Exit Sub
    Next
End Sub

' -----------------------------------------------------------------------------
' TryDisplayCurve (Function), source lines 159-171
' -----------------------------------------------------------------------------
Private Function TryDisplayCurve(ByRef shape As VGCore.shape) As Boolean
    On Error GoTo 10
    LengthOfCurves = LengthOfCurves + Length(shape.DisplayCurve)
    AreaOfCurves = AreaOfCurves + Area(shape.DisplayCurve)
    RectAreaOfCurves = RectAreaOfCurves + RectArea(shape.DisplayCurve)
    
    CountSubPaths = CountSubPaths + shape.DisplayCurve.Subpaths.count
    CountNodes = CountNodes + shape.DisplayCurve.Nodes.count
    
    TryDisplayCurve = True
    Exit Function
10
End Function

' -----------------------------------------------------------------------------
' CountErrPlus1 (Sub), source lines 172-191
' -----------------------------------------------------------------------------
Private Sub CountErrPlus1(ByRef shape As VGCore.shape)
    CountErr = CountErr + 1
    ReDim Preserve ErrOnScreenCurves(CountErr)
    ReDim Preserve ErrOnScreenText(CountErr)
    
    Set ErrOnScreenText(CountErr) = New OnScreenText
    ErrOnScreenText(CountErr).SetTextColor RGB(250, 0, 0)
    ErrOnScreenText(CountErr).SetPixelOffset 15, -10
    ErrOnScreenText(CountErr).SetTextAndPosition "ERR-" & CountErr, shape.PositionX, shape.PositionY, 2 'cdrCenterOnScreenText ' cdrCenterOnSreenText
    'ErrOnScreenText(CountErr).Show

    Set ErrOnScreenCurves(CountErr) = New OnScreenCurve
    With shape.boundingBox
        ErrOnScreenCurves(CountErr).SetPen RGB(250, 0, 0), 1, cdrOnScreenCurvePenSolid
        Dim curve As curve
        Set curve = CreateErrCurve(.x, .y, .right, .top)
        ErrOnScreenCurves(CountErr).SetCurve curve
        'ErrOnScreenCurves(CountErr).Show
    End With
End Sub

' -----------------------------------------------------------------------------
' CreateErrCurve (Function), source lines 192-230
' -----------------------------------------------------------------------------
Private Function CreateErrCurve(x1 As Double, y1 As Double, x2 As Double, y2 As Double) As curve
    Set CreateErrCurve = New curve
    Dim ce(1 To 9) As CurveElement
    ce(1).ElementType = cdrElementStart
    ce(1).PositionX = x1
    ce(1).PositionY = y1
    ce(2).ElementType = cdrElementLine
    ce(2).PositionX = x2
    ce(2).PositionY = y2
    
    ce(3).ElementType = cdrElementStart
    ce(3).PositionX = x1
    ce(3).PositionY = y2
    ce(4).ElementType = cdrElementLine
    ce(4).PositionX = x2
    ce(4).PositionY = y1
  '===================================
    ce(5).ElementType = cdrElementStart
    ce(5).PositionX = x1
    ce(5).PositionY = y2
    
    ce(6).ElementType = cdrElementLine
    ce(6).PositionX = x2
    ce(6).PositionY = y2
    
    ce(7).ElementType = cdrElementLine
    ce(7).PositionX = x2
    ce(7).PositionY = y1
    
    ce(8).ElementType = cdrElementLine
    ce(8).PositionX = x1
    ce(8).PositionY = y1
    
    ce(9).ElementType = cdrElementLine
    ce(9).PositionX = x1
    ce(9).PositionY = y2
    
    CreateErrCurve.CreateSubPathFromArray ce
End Function

' -----------------------------------------------------------------------------
' Length (Function), source lines 232-237
' -----------------------------------------------------------------------------
Private Function Length(ByRef crv As VGCore.curve) As Double
    Dim q As Long
     For q = 1 To crv.Subpaths.count
        Length = Length + crv.Subpaths(q).Length
     Next
End Function

' -----------------------------------------------------------------------------
' Area (Function), source lines 238-268
' -----------------------------------------------------------------------------
Private Function Area(ByRef crv As VGCore.curve) As Double
    Dim q As Long, q2 As Long
    Dim IsInner() As Boolean
    ReDim IsInner(crv.Subpaths.count)

    For q = 1 To crv.Subpaths.count
        With crv.Subpaths(q)
            For q2 = 1 To crv.Subpaths.count
                If q <> q2 Then
                    If crv.Subpaths(q2).IsOnSubPath(.Nodes.First.PositionX, .Nodes.First.PositionY, 0) = cdrInsideShape Then IsInner(q) = Not IsInner(q)
                End If
            Next
        End With
        
        If Timer - t > 0.02 Then
            t = Timer
            If PerimetrFormVisible = True Then PerimetrForm.UserFormRefreshData
            If SquareFormVisible = True Then
                SquareForm.UserFormRefreshData
                If SquareForm.ComboBox1.ListIndex = 1 Then ShowRectOnScreenCurves Else HideRectOnScreenCurves
            End If
            DoEvents
        End If
        If NeedBreak = True Then Exit Function
    Next
    
    For q = 1 To crv.Subpaths.count
        If IsInner(q) = False Then Area = Area + crv.Subpaths(q).Area Else Area = Area - crv.Subpaths(q).Area
    Next
    
End Function

' -----------------------------------------------------------------------------
' RectArea (Function), source lines 269-321
' -----------------------------------------------------------------------------
Private Function RectArea(ByRef crv As VGCore.curve) As Double
    Dim px() As Double, Py() As Double
    Dim rX() As Double, dy() As Double
    Dim deleted() As Boolean
    ReDim px(crv.Subpaths.count), Py(crv.Subpaths.count)
    ReDim rX(crv.Subpaths.count), dy(crv.Subpaths.count)
    ReDim deleted(crv.Subpaths.count)
    Dim q As Long, q2 As Long
    
    For q = 1 To crv.Subpaths.count
        px(q) = crv.Subpaths(q).boundingBox.left
    Next
    
    For q = 1 To crv.Subpaths.count
        Py(q) = crv.Subpaths(q).boundingBox.top
    Next
    
    For q = 1 To crv.Subpaths.count
        rX(q) = crv.Subpaths(q).boundingBox.right
    Next
    
    For q = 1 To crv.Subpaths.count
        dy(q) = crv.Subpaths(q).boundingBox.bottom
    Next

    For q = 1 To crv.Subpaths.count
        For q2 = 1 To crv.Subpaths.count
            If q <> q2 Then
                If px(q) > px(q2) Then
                    If rX(q) < rX(q2) Then
                        If Py(q) < Py(q2) Then
                            If dy(q) > dy(q2) Then
                                deleted(q) = True
                                Exit For
                            End If
                        End If
                    End If
                End If
            End If
        Next
    Next
    
    For q = 1 To crv.Subpaths.count
        If deleted(q) = False Then
            RectArea = RectArea + (rX(q) - px(q)) * (Py(q) - dy(q))
            RectOnScreenCurvesCount = RectOnScreenCurvesCount + 1
            ReDim Preserve RectOnScreenCurves(RectOnScreenCurvesCount)
            Set RectOnScreenCurves(RectOnScreenCurvesCount) = New OnScreenCurve
            RectOnScreenCurves(RectOnScreenCurvesCount).SetRectangle px(q), Py(q), rX(q), dy(q)
            RectOnScreenCurves(RectOnScreenCurvesCount).SetBrush RGB(230, 230, 255)
        End If
    Next
End Function
