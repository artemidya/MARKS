Private Sub MarginCheck()
    Select Case lZeroPosition
        Case RightP
            If lXmargin < 0 Then lXmargin = 0
            If lYmargin < 0 Then lYmargin = 0
            If lXmargin > MaxSheetHeight Then lXmargin = MaxSheetHeight
            If lYmargin > MaxSheetWidth Then lYmargin = MaxSheetWidth
        Case LeftP
            If lXmargin < 0 Then lXmargin = 0
            If lYmargin < 0 Then lYmargin = 0
            If lXmargin > MaxSheetWidth Then lXmargin = MaxSheetWidth
            If lYmargin > MaxSheetHeight Then lYmargin = MaxSheetHeight
        Case CenterP
            If lXmargin < -MaxSheetWidth / 2 Then lXmargin = -MaxSheetWidth / 2
            If lYmargin < -MaxSheetHeight / 2 Then lYmargin = -MaxSheetHeight / 2
            If lXmargin > MaxSheetWidth / 2 Then lXmargin = MaxSheetWidth / 2
            If lYmargin > MaxSheetHeight / 2 Then lYmargin = MaxSheetHeight / 2
    End Select
End Sub
