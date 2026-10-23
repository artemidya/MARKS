Private Function GetStandardCarriageCurves() As Collection
    If Not StandardCarriageCurvesInitialized Then InitStandardCarriage
    Set GetStandardCarriageCurves = StandardCarriageCurves
End Function
