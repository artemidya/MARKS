Private Sub ComboBox2_Change()
    With ComboBox2
        If .ListIndex > -1 Then
            Profiles(TabStrip1.SelectedItem.index).ActiveSpeed = .ListIndex
            ComboBox6.ListIndex = .ListIndex
        End If
    End With
End Sub
