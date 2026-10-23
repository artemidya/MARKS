Private Function CloseCurve(crv As VGCore.curve) As VGCore.curve
    If crv.closed Then
        Set CloseCurve = crv
    Else
        Set CloseCurve = VGCore.ActiveDocument.CreateCurve
        CloseCurve.AppendCurve crv
        CloseCurve.closed = True
    End If
End Function
