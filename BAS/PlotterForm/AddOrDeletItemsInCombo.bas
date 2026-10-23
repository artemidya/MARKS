Private Sub AddOrDeletItemsInCombo(n As Long, Combo As ComboBox)
    Dim q As Long
    If n < 0 Then
        For q = 1 To Abs(n)
            Combo.RemoveItem (Combo.ListCount - 1)
        Next
    Else
        For q = 1 To n
            Combo.AddItem ""
        Next
    End If
End Sub
