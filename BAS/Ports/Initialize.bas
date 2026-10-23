Private Sub Initialize(i As Long)
    On Error GoTo 10
    With Ports(i)
'        .profileIndex = PlotterForm.TabStrip1.SelectedItem.index
        .Init = Profiles(PlotterForm.TabStrip1.SelectedItem.index).Initializer
        .TERM = Profiles(PlotterForm.TabStrip1.SelectedItem.index).Terminator
        .Initialized = True
        
        PlotterForm.AddLog .Init
        Dim bytes() As Byte
        bytes = StrToAscii(.Init)
        
        If UBound(bytes) > -1 Then
            COM.CommOpen 1, Trim(.name), RemoveTrailingColon(Trim(.name))
            COM.CommClose 1
            Dim fileNumber As Integer: fileNumber = FreeFile
            Open Trim(.name) For Binary Access Read Write As #fileNumber
                Put #fileNumber, , bytes
            Close #fileNumber
        End If
    End With
10
End Sub
