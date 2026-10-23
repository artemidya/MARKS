Public Function ListIndexByName(name As String, Combo As ComboBox, Optional defIndex As Long = 0) As Long
    Dim q As Long
    With Combo
        For q = 0 To .ListCount - 1
            If KillBirdsAndSpaces(.List(q)) = name Then
                ListIndexByName = q
                Exit Function
            End If
        Next
        If defIndex < .ListCount Then
            ListIndexByName = defIndex 'Select deafault
        Else
            ListIndexByName = 0
        End If
    End With
End Function
