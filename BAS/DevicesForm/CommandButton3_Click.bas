Private Sub CommandButton3_Click()
    CommandButton3.Enabled = False
    DataFromMeToProfile
    Profiles(PlotterForm.TabStrip1.SelectedItem.index).Save
    
    If DeveloperMode Then
        If Authorization Then
            'Send
            Dim content As String
            
            If Len(Trim(Profiles(PlotterForm.TabStrip1.SelectedItem.index).PlotterModelString)) > 10 Then
                'Creating a temporary object for converting equipment image to SVG
                Dim PltDocument1 As New PltDocument
                With PltDocument1
                    .MaxSheetHeight = Profiles(PlotterForm.TabStrip1.SelectedItem.index).MaxSheetHeight * 1000
                    .MaxSheetWidth = Profiles(PlotterForm.TabStrip1.SelectedItem.index).MaxSheetWidth * 1000
                    .PlotterModelString = Profiles(PlotterForm.TabStrip1.SelectedItem.index).PlotterModelString
                    .PlotterName = Profiles(PlotterForm.TabStrip1.SelectedItem.index).Device
                    .ZeroPosition = Profiles(PlotterForm.TabStrip1.SelectedItem.index).ZeroPosition
                    content = .SVG_Generator
                End With
                
                Dim buffer2 As New buffer
                With buffer2
                    .InterPlatformNetConverter = True
                    .size = Len(content) * 2 + 4
                    .bPut_Long Commands.SetSVG
                    .bPut_String content
                    BlocksSocket1.Send .bytes
                End With
            End If
            
            content = Profiles(PlotterForm.TabStrip1.SelectedItem.index).content
            Dim buffer1 As New buffer
            With buffer1
                .InterPlatformNetConverter = True
                .size = Len(content) * 2 + 4
                .bPut_Long Commands.SetPlotter
                .bPut_String content
                BlocksSocket1.Send .bytes
            End With
            
        End If
    End If
End Sub
