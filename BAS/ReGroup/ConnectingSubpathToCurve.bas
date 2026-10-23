Private Function ConnectingSubpathToCurve(SubPath As VGCore.SubPath, crv As VGCore.curve) As Boolean
    Dim q As Long
    With crv
1:
        For q = 1 To .Subpaths.count
            If Not .Subpaths(q).closed Then
                If .Subpaths(q).segments.count > 0 Then
                    If ConnectTwoSubpathsFromDifferentCurves(SubPath, .Subpaths(q)) Then
                        ConnectingSubpathToCurve = True
                        If .segments.count = 0 Then
                            Exit Function
                        Else
                            GoTo 1
                        End If
                    End If
                End If
            End If
        Next
    End With
End Function
