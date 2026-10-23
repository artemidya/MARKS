Private Sub Text1_Change()
    If ActiveSelection.shapes.count > 0 Then
        If ConvertLng(Text1.text) > 0 And ConvertLng(Text1.text) < (1000 \ ActiveSelection.shapes.count + 1) Then
            Command4.Enabled = True
        Else
            Command4.Enabled = False
        End If
    End If
End Sub
