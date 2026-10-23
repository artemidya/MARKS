Private Sub TryConvertToCurves(shape As VGCore.shape)
    Application.Optimization = True
    Dim s2 As VGCore.shape
    Set s2 = shape.Clone
    On Error GoTo 10
    s2.ConvertToCurves
    If s2.Type = cdrCurveShape Then
        OutString = OutString & GetColor(s2)
        GetCurve s2.DisplayCurve
    ElseIf s2.Type = cdrGroupShape Then
        GetCurvesCycle s2.shapes
    End If
10
    s2.Delete
    Application.Optimization = False
End Sub
