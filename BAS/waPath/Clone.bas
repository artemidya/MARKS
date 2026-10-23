Public Function Clone() As waPath
    Dim cPtr As LongPtr
    GdipClonePath Ptr, cPtr
    
    Set Clone = New waPath
    Dim tmpPtr As LongPtr
    
    tmpPtr = Clone.Ptr
    Clone.Ptr = cPtr
    GdipDeletePath tmpPtr
    
    Set Clone.brush = lBrush
    Set Clone.pen = lPen
    

End Function
