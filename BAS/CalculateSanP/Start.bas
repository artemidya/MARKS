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
