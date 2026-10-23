Private Sub GetCurvesCycle(shapes As VGCore.shapes, Optional UpLevel As Boolean)
    Dim q1 As Long
    For q1 = shapes.count To 1 Step -1
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
                
                
                OutString = OutString & GetColor(shapes(q1))
                GetCurve .DisplayCurve
            ElseIf .Type = cdrGroupShape Or .Type = cdrContourGroupShape Or .Type = cdrArtisticMediaGroupShape Then
                GetCurvesCycle .shapes
            ElseIf TryDisplayCurve(shapes(q1)) Then
                If .Type <> cdrBitmapShape Then
                    OutString = OutString & GetColor(shapes(q1))
                    GetCurve .DisplayCurve
                End If
            Else
                'for text
                TryConvertToCurves shapes(q1)
            End If
        End With
    Next
End Sub
