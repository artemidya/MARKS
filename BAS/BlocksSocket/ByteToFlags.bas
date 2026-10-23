Private Static Function ByteToFlags(ByVal b As Byte) As SocketSystemFlags
    With ByteToFlags
        .NeedReplaceFlag = b \ 128
        .CompressFlag = b \ 64 And 1
        .CryptoFlag = b \ 32 And 1
        .PartFlag = b \ 16 And 1
        .SendFilesFlag = b \ 8 And 1
        .SystemFlag = b \ 4 And 1
        If (b And 3) = 0 Then
            .Denominator_2bit = 16777212
        ElseIf (b And 3) = 1 Then
            .Denominator_2bit = 4194304
        ElseIf (b And 3) = 2 Then
            .Denominator_2bit = 65536
        Else 'If(b And 3) = 3
            .Denominator_2bit = 1024
        End If
    End With
End Function
