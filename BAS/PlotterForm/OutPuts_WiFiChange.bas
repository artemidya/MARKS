Private Sub OutPuts_WiFiChange()
    Dim q As Long
    q = Combo1.ListCount - 2

    With Combo1
       
        If OutPuts.Connected(q) = True Then
            If .List(q) <> ChrW(10003) & OutPuts.Names(q) Then
                suppressChange = True
                .List(q) = ChrW(10003) & OutPuts.Names(q)
                suppressChange = False
                Combo1_Change
            End If
        Else
            If .List(q) <> "    " & OutPuts.Names(q) Then
                suppressChange = True
                .List(q) = "    " & OutPuts.Names(q)
                suppressChange = False
                Combo1_Change
            End If
        End If
       
    End With
End Sub
