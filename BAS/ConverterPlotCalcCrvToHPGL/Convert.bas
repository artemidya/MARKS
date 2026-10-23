Public Function Convert(Curves As Collection, Profile As Profile, ByVal dx As Long, ByVal dy As Long, ByVal AllCurvesHeight As Long, ByVal AllCurvesWidth As Long) As String
    Last_X = -2147483647
    Last_Y = -2147483647
    LastCommandIsPD = False
    
    Set MyCurves = CloneCollection(Curves)
    Dim path As waPath
    Dim path2 As waPath
    Dim mx As New waMatrix
            
    'LastAngel =
    Select Case Profile.ZeroPosition
        Case ZeroPositions.RightP
            mx.Rotate 270
            mx.Translate 0, AllCurvesWidth + 1
            If Profile.KnifeOffset > 0 Then
                'Add KnifeOffset starting line if need
                Set path = New waPath
                path.StartSubPath
                path.AddLineSegment AllCurvesWidth, 0
                path.AddLineSegment AllCurvesWidth - Profile.KnifeOffset * 3, 0
                MyCurves.Add path, , 1
                
                Set path2 = New waPath
                path2.StartSubPath
                path2.AddLineSegment AllCurvesWidth, AllCurvesHeight + 1000
                path2.AddLineSegment AllCurvesWidth - Profile.KnifeOffset * 3, AllCurvesHeight + 1000
                MyCurves.Add path2
            End If
            LastAngel = 3.14159265358979
        Case ZeroPositions.CenterP
            mx.Translate -AllCurvesWidth / 2, -AllCurvesHeight / 2
            If Profile.KnifeOffset > 0 Then
                'Add KnifeOffset starting line if need
                Set path = New waPath
                path.StartSubPath
                path.AddLineSegment 0, 0
                path.AddLineSegment Profile.KnifeOffset * 3, 0
                MyCurves.Add path, , 1
                
                Set path2 = New waPath
                path2.StartSubPath
                path2.AddLineSegment 0, AllCurvesHeight + 1000
                path2.AddLineSegment Profile.KnifeOffset * 3, AllCurvesHeight + 1000
                MyCurves.Add path2
            End If
            LastAngel = 1.5707963267949
        Case ZeroPositions.LeftP
            If Profile.KnifeOffset > 0 Then
                'Add KnifeOffset starting line if need
                Set path = New waPath
                path.StartSubPath
                path.AddLineSegment 0, 0
                path.AddLineSegment Profile.KnifeOffset * 3, 0
                MyCurves.Add path, , 1
                
                Set path2 = New waPath
                path2.StartSubPath
                path2.AddLineSegment 0, AllCurvesHeight + 1000
                path2.AddLineSegment Profile.KnifeOffset * 3, AllCurvesHeight + 1000
                MyCurves.Add path2
            End If
            LastAngel = 1.5707963267949
    End Select
    
    'Apply Matrix
    For Each path In MyCurves
        path.ApplyMatrixTransform mx
    Next

    Convert = CurvesToString(dx, dy, Profile)
End Function
