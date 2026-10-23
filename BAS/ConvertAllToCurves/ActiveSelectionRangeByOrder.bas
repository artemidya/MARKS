Private Function ActiveSelectionRangeByOrder() As VGCore.ShapeRange
    Dim q2 As Long
    Dim newSr As New VGCore.ShapeRange
    With ActiveDocument.ActivePage
        For q2 = 1 To .shapes.count
            If .shapes(q2).Selected Then
                newSr.Add .shapes(q2)
            End If
        Next
    End With
    Set ActiveSelectionRangeByOrder = newSr.ReverseRange
End Function
