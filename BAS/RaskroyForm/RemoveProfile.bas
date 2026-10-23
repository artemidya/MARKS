Private Sub RemoveProfile(ByVal i As Long)
    UserData.CountRaskroyProfiles = UserData.CountRaskroyProfiles - 1
    Dim q As Long
    For q = i To UserData.CountRaskroyProfiles - 1
        UserData.RaskroyProfiles(q) = UserData.RaskroyProfiles(q + 1)
    Next
    If UserData.CountRaskroyProfiles > 0 Then
        ReDim Preserve UserData.RaskroyProfiles(UserData.CountRaskroyProfiles - 1)
    Else
        Erase UserData.RaskroyProfiles
    End If
    UserData.Save
End Sub
