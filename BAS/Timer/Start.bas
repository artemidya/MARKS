Public Sub Start(milliseconds As Long)
    Interval = milliseconds
    Dim dblEnd As Single
    Do While Processing
        Sleep 10
        DoEvents
        Do While Enabled
            dblEnd = Timer + Interval / 1000
            Do While Timer < dblEnd And Enabled
                Sleep 10
                DoEvents
            Loop
            If Enabled Then RaiseEvent OnTimer
        Loop
    Loop
End Sub
