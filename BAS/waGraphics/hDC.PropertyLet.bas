Property Let hDC(oValue As LongPtr)
1
    If Not Initialized Then
        lhDC = oValue
        Init
    Else
        Class_Terminate
        GoTo 1
    End If
End Property
