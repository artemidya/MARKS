Private Sub AddToStatisticSent(ByVal cntBytes As Long)
    'Inline
    Statistic_BytesSent = Statistic_BytesSent + cntBytes
    If zProtected_MeHaweTheParentFlag Then
        With Parent
            .Statistic_BytesSent = .Statistic_BytesSent + cntBytes
        End With
    End If
    zProtected_RaiseEvent_OnStatisticChange Me
End Sub
