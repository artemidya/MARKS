Public Sub ShowRectOnScreenCurves()
    On Error GoTo 10
    For CurQ_ShowRectOnScreenCurves = CurQ_ShowRectOnScreenCurves To RectOnScreenCurvesCount
        RectOnScreenCurves(CurQ_ShowRectOnScreenCurves).Show
    Next
10
End Sub
