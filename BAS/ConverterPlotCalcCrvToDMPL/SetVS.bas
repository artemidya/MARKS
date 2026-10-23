Private Function SetVS(vs As Long) As String
    Static lastVS As Long
    If vs <> lastVS Then
        SetVS = "V" & Int(vs / 10) & ","
        lastVS = vs
    End If
End Function
