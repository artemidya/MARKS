Private Sub RemoveCoincidentCurvesSelfy(Curves() As VGCore.curve)
    Dim txt As String
    txt = localization.ActiveLng.Item("RemoveDuplicateLines.Command") & " (2 / 3)"
    ProgressForm.StageName = txt
    
    Dim q As Long
    For q = 0 To UBound(Curves)
        If OutlineUse Then Sensitivity = OutlineWidths(q)
        RemoveCoincidentCurveSelfy Curves(q)
        If Aborted Then Exit Sub
    Next
End Sub
