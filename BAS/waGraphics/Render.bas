Public Sub Render(Optional Draw As Boolean = True, Optional Fill As Boolean = False, Optional ShadowBrush As waBrush, Optional SpecialPen As waPen)
    If Not Curves Is Nothing Then
        Dim curve As waPath
        If ShadowBrush Is Nothing Then
            If SpecialPen Is Nothing Then
                For Each curve In Curves
                    With curve
                        If Fill Then GdipFillPath lHGraphics, .brush.Ptr, .Ptr
                        If Draw Then GdipDrawPath lHGraphics, .pen.Ptr, .Ptr
                    End With
                Next
            Else
                For Each curve In Curves
                    With curve
                        If Fill Then GdipFillPath lHGraphics, .brush.Ptr, .Ptr
                        If Draw Then GdipDrawPath lHGraphics, SpecialPen.Ptr, .Ptr
                    End With
                Next
            End If
        Else
            'Shadow
            For Each curve In Curves
                With curve
                    GdipFillPath lHGraphics, ShadowBrush.Ptr, .Ptr
                End With
            Next
        End If
    End If
End Sub
