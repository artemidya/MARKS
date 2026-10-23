Public Function GetPlotCalcCrv(ByRef StringL As String) As Collection
    On Error GoTo 10
    Set GetPlotCalcCrv = New Collection
    LenStringL = Len(StringL)
    cp = 1
    Dim brush As waBrush
    Dim path As waPath
    Do While cp <= LenStringL
        '1 GetBrush
        Set brush = GetBrush(StringL)
        '2 GetPath
        Set path = GetPath(StringL)
        '3 Add to Collection
        
        If Not path Is Nothing Then
            Set path.brush = brush
            GetPlotCalcCrv.Add path
        End If
        cp = cp + 1
    Loop
10
End Function
