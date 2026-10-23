Private Function NewZamTerminator() As Long
    'Delaet noviy zamenitel dla terminatora (0.0.0.0)
1:
    On Error GoTo 10
    Static Z As Long
    If Z = 0 Then
        Randomize Timer
        Z = CLng(Rnd * 2147483647)
    End If
    If Z < 2147483647 Then
        Z = Z + 1
    Else
        Z = -2147483647 - 1
    End If
    NewZamTerminator = Z
    Exit Function
10
      GoTo 1
End Function
