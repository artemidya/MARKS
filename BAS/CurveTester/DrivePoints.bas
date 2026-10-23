Private Sub DrivePoints(ByRef XS() As Double, ByRef YS() As Double, ByVal count As Long)
    Dim q As Long
    Dim n As Double
    Dim s1 As shape
    Dim crv As VGCore.curve
    
    Set crv = Application.CreateCurve(ActiveDocument)
    For q = 1 To count
        n = 0.2
        crv.CreateSubPath(XS(q) - n, YS(q) - n).AppendLineSegment XS(q) + n, YS(q) + n
        crv.CreateSubPath(XS(q) - n, YS(q) + n).AppendLineSegment XS(q) + n, YS(q) - n
    Next
    
    Set s1 = Application.ActiveLayer.CreateCurve(crv) '.Outline.SetProperties(Color:=CreateCMYKColor(0, 60, 100, 0))
    s1.Outline.SetProperties Color:=CreateCMYKColor(0, 100, 100, 0)
    's1.Outline.SetProperties 0.04222, OutlineStyles(0), CreateCMYKColor(0, 100, 100, 0), ArrowHeads(0), ArrowHeads(0), False, False, cdrOutlineButtLineCaps, cdrOutlineMiterLineJoin, 0#, 100

End Sub
