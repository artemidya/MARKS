Public Sub SelectLangByCode(LngCode As String)
    Select Case UCase(LngCode)
        Case "RU"
            RU
        Case "EN"
            EN
        Case "UK"
            UK
        Case "ZH"
            ZH
        Case "PL"
            PL
        Case "DE"
            DE
        Case "FR"
            FR
        Case "SK"
            SK
        Case "SL"
            SL
        Case "IT"
            IT
        Case "VI"
            VI
        Case "KK"
            KK
        Case "PT"
            PT
        Case "ES"
            ES
        Case "TR"
            TR
        Case "ID"
            ID
        Case "AR"
            AR
        Case "HI"
            HI
        Case Else
            EN
    End Select
End Sub
