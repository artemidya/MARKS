Private Sub CheckBox1_Change()
    If Not dntChangeCheckBox1 Then
        RegData1.NestingPlaceOnPage = CheckBox1.value
    End If
End Sub
