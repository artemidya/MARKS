Private Sub StartR(ShapesArray() As VGCore.shape)
    Dim q As Long
    Dim count As Long

    Dim txt As String
    txt = localization.ActiveLng.Item("ReGroup.Command")
    
    q = UBound(ShapesArray()) - 1
    count = q + 2

    Do While q > -1
        If OneOfSubpathsIsClosed(ShapesArray(q)) Then
            '------Progress-Bar-------
            If ProgressFormIsShowed Then
                If Timer - LastRefreshTime > 0.1 Then
                    ProgressForm.Percent = (count - q) / count
                    LastRefreshTime = Timer
                    ProgressForm.StageName = txt & " (" & count - q & " / " & count & ") "
                    DoEvents
                    If ProgressForm.Break Then Exit Sub
                End If
            Else
                If Timer - StartTime > 0.5 Then
                    ProgressFormIsShowed = True
                    ProgressForm.Show
                End If
            End If
            '------Progress-Bar-------
        
            ToGroup ShapesArray(), q
        End If
        q = q - 1
    Loop
End Sub
