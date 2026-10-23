Private Function ConnectingCurveToCurve(crv1 As VGCore.curve, crv2 As VGCore.curve) As Boolean
    Dim q As Long
    With crv1
        For q = 1 To .Subpaths.count
            If Not .Subpaths(q).closed Then
                If ConnectingSubpathToCurve(.Subpaths(q), crv2) Then
                    ConnectingCurveToCurve = True
                    If crv2.segments.count = 0 Then Exit Function
                End If
            End If
        Next
    End With
End Function
