Public Function GetLogoCurves() As Collection
    If Not LogoCurvesInitialized Then InitLogoCurves
    Set GetLogoCurves = LogoCurves
End Function
