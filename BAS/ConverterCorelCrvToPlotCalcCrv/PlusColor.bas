Private Sub PlusColor(Color As RGBA)
    Dim q As Long
    Dim colorInfo1 As ColorInfo

    For Each colorInfo1 In MyColors
        If colorInfo1.R = Color.R And colorInfo1.g = Color.g And colorInfo1.b = Color.b And colorInfo1.a = Color.a Then
            colorInfo1.count = colorInfo1.count + 1
            Exit Sub
        End If
    Next

    'Add color
    Set colorInfo1 = New ColorInfo
    colorInfo1.R = Color.R
    colorInfo1.g = Color.g
    colorInfo1.b = Color.b
    colorInfo1.a = Color.a
    colorInfo1.count = 1
    MyColors.Add colorInfo1
End Sub
