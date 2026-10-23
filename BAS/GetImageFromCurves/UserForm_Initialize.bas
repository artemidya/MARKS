Private Sub UserForm_Initialize()
    If ActiveSelection.shapes.count < 1 Then MsgBox "Select shape", vbInformation: End
    GetCurves Application.ActiveSelectionRange.shapes
End Sub
