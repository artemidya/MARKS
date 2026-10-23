Public Sub SortProfiles()
    Dim i As Long
    Dim fSwitched As Boolean
    Dim raskroyProfileTemp As RaskroyProfile
        Do
            fSwitched = False
            For i = 0 To CountRaskroyProfiles - 2
                If addZerroForDigits(UCase(RaskroyProfiles(i).name)) > addZerroForDigits(UCase(RaskroyProfiles(i + 1).name)) Then
                    fSwitched = True
                    raskroyProfileTemp = RaskroyProfiles(i)
                    RaskroyProfiles(i) = RaskroyProfiles(i + 1)
                    RaskroyProfiles(i + 1) = raskroyProfileTemp
                    If i = LastProfileSelectionNumber Then
                        LastProfileSelectionNumber = i + 1
                    ElseIf i + 1 = LastProfileSelectionNumber Then
                        LastProfileSelectionNumber = i
                    End If
                End If
            Next
        Loop While fSwitched
End Sub
