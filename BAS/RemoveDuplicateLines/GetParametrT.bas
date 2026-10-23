Private Function GetParametrT(x As Double, y As Double, segment As VGCore.segment) As Double
    Dim t As Double
    segment.FindClosestPoint x, y, t
    t = Round(t, 10)
    GetParametrT = t
End Function
