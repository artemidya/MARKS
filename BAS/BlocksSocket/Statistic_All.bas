Public Property Get Statistic_All(Optional ShowRcvFlags As Boolean = False) As String
    Dim q As Long, Percent As Long: Percent = Statistic_PercentReceiving
    Dim PBar As String ' [+++______________]
    Dim oldRcv As String
    PBar = "["
    For q = 1 To Percent \ 5
        PBar = PBar & "+"
    Next
    For q = q To 20
        PBar = PBar & "_"
    Next
    PBar = PBar & "] "
    If OnReciveOld_TurnON Then oldRcv = "; ReceivedOld=" & Statistic_BytesReceivedOld
    Dim sum As String
    If LengthBIGDatagramm > 0 Then
        sum = CP_BIGDatagramm & "/" & LengthBIGDatagramm
    Else
        sum = CP_Datagramm & "/" & LengthDatagramm
    End If
    Statistic_All = Address & " Sent=" & Statistic_BytesSent _
                            & "; Received=" & Statistic_BytesReceived _
                            & oldRcv _
                            & "; Skipped=" & Statistic_BytesSkipped _
                            & "; ProgressReceiving=" & PBar & Percent & "% " _
                            & sum _
                            & "; PING=" & Statistic_PING & " ms"
    If Not (Clients Is Nothing) Then
        Statistic_All = Statistic_All & vbCrLf & "==================================================" & vbCrLf
        Statistic_All = Statistic_All & "Clients.Count=" & Clients.count & vbCrLf
        Statistic_All = Statistic_All & "==================================================" & vbCrLf
        Dim client As BlocksSocket
        For Each client In Clients
             Statistic_All = Statistic_All & client.Statistic_All(ShowRcvFlags) & vbCrLf
        Next
    Else
        If ShowRcvFlags Then
            With RcvFlags
                Statistic_All = Statistic_All & vbCrLf & "RcvFlags> NeedReplaceFlag=" & .NeedReplaceFlag _
                                                       & "; CompressFlag=" & .CompressFlag _
                                                       & "; CryptoFlag=" & .CryptoFlag _
                                                       & "; PartFlag=" & .PartFlag _
                                                       & "; SendFilesFlag=" & .SendFilesFlag _
                                                       & "; SystemFlag=" & .SendFilesFlag _
                                                       & "; Denominator=" & .Denominator_2bit
            End With
            With SndFlags
                Statistic_All = Statistic_All & vbCrLf & "SndFlags> NeedReplaceFlag=" & .NeedReplaceFlag _
                                                       & "; CompressFlag=" & .CompressFlag _
                                                       & "; CryptoFlag=" & .CryptoFlag _
                                                       & "; PartFlag=" & .PartFlag _
                                                       & "; SendFilesFlag=" & .SendFilesFlag _
                                                       & "; SystemFlag=" & .SendFilesFlag _
                                                       & "; Denominator=" & .Denominator_2bit & vbCrLf
            End With
        End If
    End If
End Property
