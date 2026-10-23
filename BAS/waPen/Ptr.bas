Property Get Ptr() As LongPtr
    If Not Initialized Then SetColor DefColor, 100
    Ptr = lPtr
End Property
