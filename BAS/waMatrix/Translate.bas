Public Sub Translate(offsetX As Long, offsetY As Long)
    GdipTranslateMatrix lPtr, CSng(offsetX), CSng(offsetY), MatrixOrderAppend
End Sub
