Private Static Function FlagsToByte(Flags As SocketSystemFlags) As Byte
    With Flags
        FlagsToByte = .NeedReplaceFlag * 128
        FlagsToByte = FlagsToByte Or .CompressFlag * 64
        FlagsToByte = FlagsToByte Or .CryptoFlag * 32
        FlagsToByte = FlagsToByte Or .PartFlag * 16
        FlagsToByte = FlagsToByte Or .SendFilesFlag * 8
        FlagsToByte = FlagsToByte Or .SystemFlag * 4
        If .Denominator_2bit = 16777212 Then
            Exit Function
        ElseIf .Denominator_2bit = 4194304 Then
            FlagsToByte = FlagsToByte Or 1
        ElseIf .Denominator_2bit = 65536 Then
            FlagsToByte = FlagsToByte Or 2
        Else    'If .Denominator_2bit = 1024
            FlagsToByte = FlagsToByte Or 3
        End If
    End With
End Function
