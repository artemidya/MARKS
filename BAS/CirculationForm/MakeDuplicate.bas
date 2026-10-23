Private Sub MakeDuplicate(n As Long)
    If n > 1 Then
        Application.Optimization = True
        Application.ActiveDocument.BeginCommandGroup localization.ActiveLng.Item("CommandBars.ToolTipText1_5") & " (" & n & ")"
        Dim q As Long
        Dim OrigSelection As ShapeRange
        Set OrigSelection = ActiveSelectionRange
        Dim dup() As ShapeRange
        ReDim dup(n - 2)
        
        For q = 0 To n - 2
            Set dup(q) = OrigSelection.Duplicate
            dup(q).Move 0, OrigSelection.SizeHeight * (q + 1)
        Next
        
        OrigSelection.AddToSelection
        For q = 0 To n - 2
            dup(q).AddToSelection
        Next
        Application.ActiveDocument.EndCommandGroup
        Application.Optimization = False
        Application.Refresh
    End If
End Sub
