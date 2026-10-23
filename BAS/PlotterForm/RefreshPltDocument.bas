Private Sub RefreshPltDocument()
    If TabStrip1.value > 7 Then Exit Sub
    If OriginalCurves Is Nothing Then Exit Sub
    If PltDocument1(TabStrip1.value) Is Nothing Then
        Me.Repaint
        Set PltDocument1(TabStrip1.value) = New PltDocument
        With PltDocument1(TabStrip1.value)
            Dim kdpi As Double
            kdpi = MonitorDPI.GetDpiDbl
            .SetScreenSize 760 * kdpi, 552 * kdpi
            .PositionX = 400 * kdpi
            .PositionY = 47 * kdpi
            
            If Profiles(TabStrip1.value).MaxSheetHeight = 0 Then
                .MaxSheetHeight = 2147483647
            Else
                .MaxSheetHeight = Profiles(TabStrip1.value).MaxSheetHeight * 1000
            End If
            
            If Profiles(TabStrip1.value).MaxSheetWidth = 0 Then
                If Profiles(TabStrip1.value).Turn = Turn0 Or Profiles(TabStrip1.value).Turn = Turn180 Then
                    If OriginalWidth <= 600000 Then
                        .MaxSheetWidth = 600000
                    ElseIf OriginalWidth <= 1600000 Then
                        .MaxSheetWidth = 1600000
                    Else
                        .MaxSheetWidth = OriginalWidth
                    End If
                Else
                    If OriginalHeight <= 600000 Then
                        .MaxSheetWidth = 600000
                    ElseIf OriginalHeight <= 1600000 Then
                        .MaxSheetWidth = 1600000
                    Else
                        .MaxSheetWidth = OriginalHeight
                    End If
                End If
            Else
                .MaxSheetWidth = Profiles(TabStrip1.value).MaxSheetWidth * 1000
            End If
          

            .PlotterModelString = Profiles(TabStrip1.value).PlotterModelString
            .PlotterName = Profiles(TabStrip1.value).Device

            

            Set .MaterialBrush = CreateBrush(RGB(MaxColor.R, MaxColor.g, MaxColor.b), CSng(MaxColor.a) / 255 * 100)
            .SetCurvesByVal OriginalCurves, OriginalHeight, OriginalWidth, Profiles(TabStrip1.value).Turn, Profiles(TabStrip1.value).HeightMirror, Profiles(TabStrip1.value).WidthMirror
            .ZeroPosition = Profiles(TabStrip1.value).ZeroPosition
            
            Dim p As POINTL
            p.x = OutPuts.Xmm * 1000
            p.y = OutPuts.Ymm * 1000
            PltDocument1(TabStrip1.value).XYmargin = p
            
            .Render
        End With
    Else
        PltDocument1(TabStrip1.value).Refresh
    End If
End Sub
