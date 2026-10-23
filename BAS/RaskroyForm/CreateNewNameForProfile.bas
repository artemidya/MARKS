Private Function CreateNewNameForProfile(Optional ByVal n As Long = 1) As String
    Dim q As Long
    For q = 0 To UserData.CountRaskroyProfiles - 1
        If n = 1 Then
            If Trim(UserData.RaskroyProfiles(q).name) = localization.ActiveLng.Item("RaskroyForm.NewProfile") Then
                n = n + 1
                CreateNewNameForProfile = CreateNewNameForProfile(n)
                Exit Function
            End If
        Else
            If Trim(UserData.RaskroyProfiles(q).name) = localization.ActiveLng.Item("RaskroyForm.NewProfile") & " " & n Then
                n = n + 1
                CreateNewNameForProfile = CreateNewNameForProfile(n)
                Exit Function
            End If
        End If
    Next
    
    If n = 1 Then
        CreateNewNameForProfile = localization.ActiveLng.Item("RaskroyForm.NewProfile")
    Else
        CreateNewNameForProfile = localization.ActiveLng.Item("RaskroyForm.NewProfile") & " " & n
    End If
End Function
