Property Let EnableSkip(s As Boolean)
    CommandButton1.Visible = s
    If EnableSkip Then
        Skip = False
    End If
End Property
