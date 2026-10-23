Private Sub GetCurvesCycle(shapes As VGCore.shapes)
    Dim q1 As Long
    For q1 = shapes.count To 1 Step -1
        DoEvents
        With shapes(q1)
            If .Type = cdrCurveShape Then
            
            
                GetCurve .curve, forBrushString(shapes(q1))
                
            ElseIf .Type = cdrGroupShape Or .Type = cdrContourGroupShape Or .Type = cdrArtisticMediaGroupShape Then
                GetCurvesCycle .shapes
            ElseIf TryDisplayCurve(shapes(q1)) Then
                If .Type <> cdrBitmapShape Then GetCurve .DisplayCurve, forBrushString(shapes(q1))
            Else
                'for text
                TryConvertToCurves shapes(q1)
            End If
        End With
    Next
End Sub
