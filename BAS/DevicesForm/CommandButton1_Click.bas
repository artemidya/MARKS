Private Sub CommandButton1_Click()
    'AddNewPlotter
    Dim Device As String
    Device = Trim((InputBox(localization.ActiveLng.Item("DevicesForm.InputBox"), "PlotCalc", Profiles(PlotterForm.TabStrip1.SelectedItem.index).ForModeratorDevice)))
    If Device <> "" Then
        Dim buffer1 As New buffer
        With buffer1
            .InterPlatformNetConverter = True
            .size = Len(Device) * 2 + 4
            .bPut_Long Commands.AddPlotter
            .bPut_String Device
            BlocksSocket1.Send .bytes
            Profiles(PlotterForm.TabStrip1.SelectedItem.index).ForModeratorDevice = Device
        End With
    End If
End Sub
