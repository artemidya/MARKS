Private Sub CommandButton4_Click()
    
    Dim buffer1 As New buffer
    If ListFocusNumber = 2 Then
        If ListBox2.ListIndex > -1 Then
            If Authorization Then
                'Reject from NeedAddList
                With buffer1
                    .InterPlatformNetConverter = True
                    .size = Len(NeedAddList(ListBox2.ListIndex).Device) * 2 + 4
                    .bPut_Long Commands.Reject
                    .bPut_String NeedAddList(ListBox2.ListIndex).Device
                    BlocksSocket1.Send .bytes
                End With
            End If
        End If
    Else
        If ListBox1.ListIndex > 1 Then
            If MsgBox("Delete this plotter?", vbYesNoCancel Or vbQuestion) = vbYes Then
                If Authorization Then
                    'Delete Plotter
                    With buffer1
                        .InterPlatformNetConverter = True
                        .size = 8
                        .bPut_Long Commands.DeletePlotter
                        .bPut_Long PlottersList(ListBox1.ListIndex - 2).ID
                        BlocksSocket1.Send .bytes
                    End With
                End If
            End If
        End If
    End If
End Sub
