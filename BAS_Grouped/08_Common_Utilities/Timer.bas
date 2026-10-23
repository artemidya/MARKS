' =============================================================================
' Group: 08_Common_Utilities
' Source module: Timer.cls
' Procedures: 3
' =============================================================================

' -----------------------------------------------------------------------------
' Class_Initialize (Sub), source lines 20-22
' -----------------------------------------------------------------------------
Private Sub Class_Initialize()
    Processing = True
End Sub

' -----------------------------------------------------------------------------
' Start (Sub), source lines 23-38
' -----------------------------------------------------------------------------
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

' -----------------------------------------------------------------------------
' StopTimer (Sub), source lines 39-42
' -----------------------------------------------------------------------------
Public Sub StopTimer()
    Enabled = False
    Processing = False
End Sub
