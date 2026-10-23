Private Sub GetCurvesCycle(shapes As VGCore.shapes, mx As waMatrix, Optional UpLevel As Boolean)
    Dim q1 As Long
    For q1 = 1 To shapes.count
        If UpLevel Then
            If Timer - lastT > RefreshTime Then
                Percent = (CSng(q1) / shapes.count) * 100
                lastT = Timer
                RaiseEvent OnPercentChange(Percent)
            End If
        End If
        DoEvents
        If Processing = False Then Exit Sub
        With shapes(q1)
            If .Type = cdrCurveShape Then
                GetColor shapes(q1)
                GetCurve .DisplayCurve, mx
            ElseIf .Type = cdrGroupShape Or .Type = cdrContourGroupShape Or .Type = cdrArtisticMediaGroupShape Then
                GetCurvesCycle .shapes, mx
            ElseIf TryDisplayCurve(shapes(q1)) Then
                If .Type <> cdrBitmapShape Then
                    GetColor shapes(q1)
                    GetCurve .DisplayCurve, mx
                End If
            Else
                'for text
                TryConvertToCurves shapes(q1), mx
            End If
        End With
    Next
End Sub
