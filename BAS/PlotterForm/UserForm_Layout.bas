Private Sub UserForm_Layout()
    If TabStrip1.value > 7 Then Exit Sub
    If Not PltDocument1(TabStrip1.value) Is Nothing Then
        PltDocument1(TabStrip1.value).Refresh
    End If
End Sub
