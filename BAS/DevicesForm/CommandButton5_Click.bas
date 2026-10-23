Private Sub CommandButton5_Click()
    Dim clipboard1 As New ClipBoard
    If ListFocusNumber = 2 Then
        If ListBox2.ListIndex > -1 Then
            clipboard1.SetData NeedAddList(ListBox2.ListIndex).Device
        End If
    Else
        If ListBox1.ListIndex > 1 Then
            clipboard1.SetData PlottersList(ListBox1.ListIndex - 2).Device
        End If
    End If
End Sub
