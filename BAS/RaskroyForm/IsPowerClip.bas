Private Function IsPowerClip(sh As shape) As Boolean
    On Error GoTo 10
    If sh.PowerClip.shapes.count > 0 Then IsPowerClip = True: Exit Function
10:
End Function
