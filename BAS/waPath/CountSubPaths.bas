Property Get CountSubPaths() As Long
    Dim iterator As LongPtr
    GdipCreatePathIter iterator, Ptr
    GdipPathIterGetSubpathCount iterator, CountSubPaths
    GdipDeletePathIter iterator
End Property
