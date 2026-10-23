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
