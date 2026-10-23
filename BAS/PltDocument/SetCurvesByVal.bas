Public Sub SetCurvesByVal(OriginalCurves As Collection, Height As Long, Width As Long, Optional Turn As Turns, Optional VerticalMirror As Boolean, Optional HorizontalMirror As Boolean)

    Set Curves = CloneCollection(OriginalCurves)
    
    If Turn <> Turn0 Or VerticalMirror Or HorizontalMirror Then
        'Need create Matrix
        Dim mx As New waMatrix
    End If
    
    If Turn = Turn0 Or Turn = Turn180 Then
        AllCurvesHeight = Height
        AllCurvesWidth = Width
    Else
        AllCurvesHeight = Width
        AllCurvesWidth = Height
    End If
    
    'Turn
    If Turn = TurnMinus90 Then
        mx.Rotate 90
        mx.Translate AllCurvesWidth, 0
    ElseIf Turn = Turn180 Then
        mx.Rotate 180
        mx.Translate AllCurvesWidth, AllCurvesHeight
    ElseIf Turn = TurnPlus90 Then
        mx.Rotate 270
        mx.Translate 0, AllCurvesHeight
    End If
    
    'Mirror
    If HorizontalMirror Then
        mx.ScaleM -1, 1
        mx.Translate AllCurvesWidth, 0
    End If
    
    If VerticalMirror Then
        mx.ScaleM 1, -1
        mx.Translate 0, AllCurvesHeight
    End If
    
    If Turn <> Turn0 Or VerticalMirror Or HorizontalMirror Then
        'Need apply Matrix
        Dim path As waPath
        For Each path In Curves
            path.ApplyMatrixTransform mx
        Next
    End If
End Sub
