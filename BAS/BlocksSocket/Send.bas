Public Sub Send(ByRef bytes() As Byte)
    If Not Not bytes Then
        'Statistic
        Dim lngBuffer() As Long
        If RcvFlags.Denominator_2bit < SndFlags.Denominator_2bit Then SndFlags.Denominator_2bit = RcvFlags.Denominator_2bit
        If UBound(bytes) + 1 > SndFlags.Denominator_2bit Then
            'System, need sent BigLengthDatagramm
            REDIM_BIG_Datagramm UBound(bytes) + 1
            SndFlags.SystemFlag = 0
            SndFlags.PartFlag = 1
            Dim sp As Long
            If Not Not SndQueueElements Then
                AddToSndQueueBytes bytes, sp, 1, FlagsToByte(SndFlags)
            Else
                Dim ostatok As Long: ostatok = (UBound(bytes) + 1) Mod SndFlags.Denominator_2bit
                Dim q As Long
                For q = 1 To (UBound(bytes) + 1) \ SndFlags.Denominator_2bit
                    lngBuffer = CreateLongPacket(bytes, SndFlags.Denominator_2bit, sp)
                    sp = sp + SndFlags.Denominator_2bit
                    If SendLng(lSocket, lngBuffer(0), (UBound(lngBuffer) + 1) * 4, 0) < 1 Then
                        AddToSndQueueLng lngBuffer
                        AddToSndQueueBytes bytes, sp, 1, FlagsToByte(SndFlags)
                        Exit Sub
                    Else
                        AddToStatisticSent (UBound(lngBuffer) + 1) * 4 - 12
                    End If
                Next
                If ostatok > 0 Then
                    lngBuffer = CreateLongPacket(bytes, ostatok, sp)
                    If SendLng(lSocket, lngBuffer(0), (UBound(lngBuffer) + 1) * 4, 0) < 1 Then
                        AddToSndQueueLng lngBuffer
                    Else
                        AddToStatisticSent (UBound(lngBuffer) + 1) * 4 - 12
                    End If
                End If
            End If
        Else
            SndFlags.SystemFlag = 0
            SndFlags.PartFlag = 0
            lngBuffer = CreateLongPacket(bytes, UBound(bytes) + 1)
            If Not Not SndQueueElements Then
                AddToSndQueueLng lngBuffer
            Else
                If SendLng(lSocket, lngBuffer(0), (UBound(lngBuffer) + 1) * 4, 0) < 1 Then
                    AddToSndQueueLng lngBuffer
                Else
                    AddToStatisticSent (UBound(lngBuffer) + 1) * 4 - 12
                End If
            End If
        End If
    End If
End Sub
