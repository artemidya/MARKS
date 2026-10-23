Private Sub FInsertNode(ByVal R As Integer)
  Dim i   As Integer
  Dim p   As Integer
  Dim cmp As Integer
  Dim Key As Integer
  Dim x   As Long
  Dim x1  As Integer
  Dim x2  As Integer
    cmp = 1
    Key = R
    p = n + 1 + textbuf(R)
    rson(R) = NIL
    lson(R) = NIL
    matchlength = 0
    Infinity& = 2147483647
    For x& = 0 To Infinity&
        If cmp >= 0 Then
            If rson(p) <> NIL Then
                p = rson(p)
             Else 'NOT RSON(P)...
                rson(p) = R
                dad(R) = p
                Exit Sub
            End If
         Else 'NOT CMP...
            If lson(p) <> NIL Then
                p = lson(p)
             Else 'NOT LSON(P)...
                lson(p) = R
                dad(R) = p
                Exit Sub
            End If
        End If
        For i = 1 To f - 1
            x1 = textbuf(R + i)
            x2 = textbuf(p + i)
            cmp = x1 - x2
            If cmp <> 0 Then
                Exit For
            End If
        Next i
        If i > matchlength Then
            matchposition = p
            matchlength = i
            If matchlength >= f Then
                Exit For
            End If
        End If
    Next x
    dad(R) = dad(p)
    lson(R) = lson(p)
    rson(R) = rson(p)
    dad(lson(p)) = R
    dad(rson(p)) = R
    If rson(dad(p)) = p Then
        rson(dad(p)) = R
     Else 'NOT RSON(DAD(P))...
        lson(dad(p)) = R
    End If
    dad(p) = NIL

End Sub
