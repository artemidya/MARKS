Private Sub ComboBox3_Change()
    With ComboBox3
        If .ListIndex > -1 Then
            If .ListIndex = 0 Then
                Profiles(TabStrip1.SelectedItem.index).ActiveForce = 0
            Else
                Profiles(TabStrip1.SelectedItem.index).ActiveForce = Profiles(TabStrip1.SelectedItem.index).Forces(.ListIndex - 1)
            End If
        End If
    End With
End Sub
