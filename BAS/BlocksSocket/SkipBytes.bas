Private Static Sub SkipBytes(ByVal BytesCount As Long)
    'inline
    BytesCount = BytesCount + RemainderCount
    RemainderCount = 0
    If LengthBIGDatagramm > 0 Then
        'SKIP BIG
        BytesCount = BytesCount + CP_BIGDatagramm
        LengthBIGDatagramm = 0
        CP_BIGDatagramm = 0
        Erase BIGDatagramm
    End If
    Statistic_BytesSkipped = Statistic_BytesSkipped + BytesCount
    If zProtected_MeHaweTheParentFlag Then
        With Parent
            .Statistic_BytesSkipped = .Statistic_BytesSkipped + BytesCount
        End With
    End If
End Sub
