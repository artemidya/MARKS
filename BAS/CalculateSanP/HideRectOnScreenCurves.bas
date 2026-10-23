Public Sub HideRectOnScreenCurves()
    On Error GoTo 10
    Dim q As Long
    For q = 1 To CurQ_ShowRectOnScreenCurves - 1
        RectOnScreenCurves(q).Hide
    Next
    CurQ_ShowRectOnScreenCurves = 1
10
End Sub
