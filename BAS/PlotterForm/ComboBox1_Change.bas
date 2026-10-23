Private Sub ComboBox1_Change()
    With ComboBox1
        If .ListIndex > -1 Then
            Profiles(TabStrip1.SelectedItem.index).ActiveAcceleration = .ListIndex
        End If
    End With
End Sub
