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
