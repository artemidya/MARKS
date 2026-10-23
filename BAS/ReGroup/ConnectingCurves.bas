Private Sub ConnectingCurves(sr As VGCore.ShapeRange)
    Dim q1 As Long, q2 As Long
    
    Dim left As Double, right As Double
    Dim top As Double, bottom As Double
    Dim crv As VGCore.curve
    Dim count As Long
    Dim txt As String
    txt = localization.ActiveLng.Item("ReGroup.ConnectingLines")
    
    q1 = 1
    With sr
        count = .count
        Do While q1 < count
            If sr(q1).Type = cdrCurveShape Then
                If OneOfSubpathsIsOpened(sr(q1).curve) Then
                    With sr(q1).boundingBox
                        left = .left - 0.00001
                        right = .right + 0.00001
                        top = .top + 0.00001
                        bottom = .bottom - 0.00001
                    End With
                    
                    '------Progress-Bar-------
                    If ProgressFormIsShowed Then
                        If Timer - LastRefreshTime > 0.1 Then
                            ProgressForm.Percent = q1 / count
                            LastRefreshTime = Timer
                            ProgressForm.StageName = txt & " (" & q1 & " / " & count & ")"
                            DoEvents
                            If ProgressForm.Break Or ProgressForm.Skip Then Exit Sub
                        End If
                    Else
                        If Timer - StartTime > 0.5 Then
                            ProgressFormIsShowed = True
                            ProgressForm.Show
                        End If
                    End If
                    '------Progress-Bar-------
                
                    q2 = q1 + 1
                    Do While q2 <= count
                        With sr(q2)
                            If .Type = cdrCurveShape Then
                                With .boundingBox
                                    If .right >= left Then
                                        If .left <= right Then
                                            If .top >= bottom Then
                                                If .bottom <= top Then
                                                    Set crv = sr(q2).curve
                                                    If ConnectingCurveToCurve(sr(q1).curve, crv) Then
                                                        
                                                        If crv.segments.count = 0 Then
                                                            sr.Remove q2
                                                            count = sr.count
                                                        End If

                                                        With sr(q1).boundingBox
                                                            left = .left - 0.00001
                                                            right = .right + 0.00001
                                                            top = .top + 0.00001
                                                            bottom = .bottom - 0.00001
                                                        End With
                                                        q2 = q1 + 1
                                                        GoTo nx:
                                                    End If
                                                End If
                                            End If
                                        End If
                                    End If
                                End With
                            End If
                        End With
                        q2 = q2 + 1
nx:
                    Loop
                End If
            End If
            q1 = q1 + 1
        Loop
    End With
End Sub
