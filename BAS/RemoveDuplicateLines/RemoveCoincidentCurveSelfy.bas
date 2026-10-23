Private Sub RemoveCoincidentCurveSelfy(curve As VGCore.curve)

    Dim left As Double
    Dim right As Double
    Dim top As Double
    Dim bottom As Double
    Dim l As Double
    
    Dim q1 As Long, q2 As Long
    Dim RCSTS As Long
    
    q1 = 1
    With curve.segments
        Do While q1 <= .count - 1
            With .Item(q1)
                With .boundingBox
                    left = .left - Sensitivity
                    right = .right + Sensitivity
                    top = .top + Sensitivity
                    bottom = .bottom - Sensitivity
                End With
                l = .Length
            End With
            
            q2 = q1 + 1
            Do While q2 <= .count
                currentOperationNumber = currentOperationNumber + 1
                If Timer - t > 0.2 Then
                    t = Timer
                    Application.Status.Progress = CLng(currentOperationNumber / OperationCount * 100) / 2
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
                
                
                If BoundingBoxCrossTest(.Item(q2).boundingBox, left, right, top, bottom) Then
                    If l >= .Item(q2).Length Then
                        q2 = q2 + RemoveCoincidentSegmentToSegment(.Item(q1), .Item(q2))
                    Else
                        RCSTS = RemoveCoincidentSegmentToSegment(.Item(q2), .Item(q1))
                        q1 = q1 + RCSTS
                        If RCSTS < 0 Then GoTo ExitLoop
                    End If
                End If
                q2 = q2 + 1
            Loop
ExitLoop:
            q1 = q1 + 1
        Loop
    End With
End Sub
