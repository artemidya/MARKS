Private Sub CopyPropertiesFrom(ShapeTo As VGCore.shape, ShapeFrom As VGCore.shape, DelFill As Boolean, DelOutline As Boolean)
    If DelFill Then
        If ShapeTo.Type = cdrCurveShape Then
            If ShapeTo.curve.closed Then
                If ShapeFrom.Outline.Type <> cdrNoOutline Then
                    If Not DelOutline Then
                        ShapeTo.CopyPropertiesFrom ShapeFrom, cdrCopyOutlineColor
                        ShapeTo.CopyPropertiesFrom ShapeFrom, cdrCopyOutlinePen
                    End If
                End If
            Else
                If Not DelOutline Or ShapeFrom.Outline.Type = cdrNoOutline Then
                    ShapeTo.CopyPropertiesFrom ShapeFrom, cdrCopyOutlineColor
                    ShapeTo.CopyPropertiesFrom ShapeFrom, cdrCopyOutlinePen
                End If
            End If
        End If
    Else
        ShapeTo.CopyPropertiesFrom ShapeFrom, cdrCopyFill
        If Not DelOutline Or ShapeFrom.Outline.Type = cdrNoOutline Then
            ShapeTo.CopyPropertiesFrom ShapeFrom, cdrCopyOutlineColor
            ShapeTo.CopyPropertiesFrom ShapeFrom, cdrCopyOutlinePen
        End If
    End If
    
End Sub
