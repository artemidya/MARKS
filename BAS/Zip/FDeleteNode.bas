Private Sub FDeleteNode(ByVal p As Integer)

  Dim q As Integer

    If dad(p) = NIL Then
        Exit Sub
    End If
    If rson(p) = NIL Then
        q = lson(p)
     ElseIf lson(p) = NIL Then 'NOT RSON(P)...
        q = rson(p)
     Else 'NOT LSON(P)...
        q = lson(p)
        If rson(q) <> NIL Then
            Do
                q = rson(q)
            Loop While rson(q) <> NIL
            rson(dad(q)) = lson(q)
            dad(lson(q)) = dad(q)
            lson(q) = lson(p)
            dad(lson(p)) = q
        End If
        rson(q) = rson(p)
        dad(rson(p)) = q
    End If
    dad(q) = dad(p)
    If rson(dad(p)) = p Then
        rson(dad(p)) = q
     Else 'NOT RSON(DAD(P))...
        lson(dad(p)) = q
    End If
    dad(p) = NIL

End Sub
