Private Sub ComboBox7_Change()
    With ComboBox7
        If .ListIndex > -1 Then
            If .ListIndex = 0 Then
                ComboBox6.ListIndex = 0
            Else
                If ComboBox6.ListIndex = 0 Then ComboBox6.ListIndex = .ListIndex
            End If
            Profiles(TabStrip1.SelectedItem.index).ActiveSpeedUP = .ListIndex
        End If
    End With
End Sub
