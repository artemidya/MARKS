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
