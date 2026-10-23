Private Function PointsFromAngel(ByVal Angel As Double, R As Long, Optional dx As Long, Optional dy As Long) As POINTL
    Angel = Angel - 1.5707963267949
    With PointsFromAngel
        .y = R * Sin(Angel) + dy
        .x = R * Cos(Angel) + dx
    End With
End Function
