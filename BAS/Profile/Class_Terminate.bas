Private Sub Class_Terminate()
    CountProfiles = CountProfiles - 1
    If NeedSave Then
        Save
    End If
End Sub
