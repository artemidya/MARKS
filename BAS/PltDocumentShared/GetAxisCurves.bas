Public Function GetAxisCurves() As Collection
    If Not AxisCurves_Initialized Then InitAxisCurves
    Set GetAxisCurves = AxisCurves
End Function
