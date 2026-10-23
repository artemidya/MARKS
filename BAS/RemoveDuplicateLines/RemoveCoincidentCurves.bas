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
