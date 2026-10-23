Private Sub ComboBox6_Change()
    With ComboBox6
        If .ListIndex > -1 Then
            If .ListIndex = 0 Then
                ComboBox7.ListIndex = 0
            Else
                If ComboBox7.ListIndex = 0 Then ComboBox7.ListIndex = .ListIndex
            End If
            Profiles(TabStrip1.SelectedItem.index).ActiveSpeed = .ListIndex
        End If
    End With
End Sub
