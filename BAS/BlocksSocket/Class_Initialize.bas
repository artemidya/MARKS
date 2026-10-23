Private Sub Class_Initialize()
    StartWinsock
    RcvFlags.Denominator_2bit = DefaultDenominator
    SndFlags.Denominator_2bit = DefaultDenominator
    Statistic_PING = -1
End Sub
