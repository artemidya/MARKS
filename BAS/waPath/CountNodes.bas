Property Get CountNodes() As Long
    GdipGetPointCount Ptr, CountNodes
'    Dim iterator As LongPtr
'    GdipCreatePathIter iterator, Ptr
'    GdipPathIterGetCount iterator, CountNodes
'    GdipDeletePathIter iterator
End Property
