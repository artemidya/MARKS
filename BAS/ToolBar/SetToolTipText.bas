Public Sub SetToolTipText()
    On Error Resume Next
    Dim RegData1 As New RegData
    Application.CommandBars("PlotCalc").Controls(1).ToolTipText = localization.ActiveLng.Item("CommandBars.ToolTipText0")
    Application.CommandBars("PlotCalc").Controls(2).ToolTipText = localization.ActiveLng.Item("CommandBars.ToolTipText1")
    Application.CommandBars("PlotCalc").Controls(3).ToolTipText = localization.ActiveLng.Item("CommandBars.ToolTipText1_5")
    Application.CommandBars("PlotCalc").Controls(4).ToolTipText = localization.ActiveLng.Item("CommandBars.ToolTipText2")
    Application.CommandBars("PlotCalc").Controls(5).ToolTipText = localization.ActiveLng.Item("CommandBars.ToolTipText3")
    Application.CommandBars("PlotCalc").Controls(6).ToolTipText = localization.ActiveLng.Item("CommandBars.ToolTipText4")
    Application.CommandBars("PlotCalc").Controls(7).ToolTipText = localization.ActiveLng.Item("CommandBars.ToolTipText5")
    Application.CommandBars("PlotCalc").Controls(8).ToolTipText = localization.ActiveLng.Item("CommandBars.ToolTipText6")
    Application.CommandBars("PlotCalc").Controls(9).ToolTipText = localization.ActiveLng.Item("CommandBars.ToolTipText7")
    Application.CommandBars("PlotCalc").Controls(10).ToolTipText = localization.ActiveLng.Item("CommandBars.ToolTipText8")
    Application.CommandBars("PlotCalc").Controls(11).ToolTipText = localization.ActiveLng.Item("CropBMP.Command") & " (" & CDbl(RegData1.CropBMPoffset) / 1000 & " " & localization.ActiveLng.Item("RaskroyForm.mm") & ")"
    Application.CommandBars("PlotCalc").Controls(12).ToolTipText = localization.ActiveLng.Item("TrimCurves.Command") & " (" & CDbl(RegData1.TrimOffset) / 1000 & " " & localization.ActiveLng.Item("RaskroyForm.mm") & ")"
    Application.CommandBars("PlotCalc").Controls(13).ToolTipText = localization.ActiveLng.Item("ReGroup.Command")
    Application.CommandBars("PlotCalc").Controls(14).ToolTipText = localization.ActiveLng.Item("RemoveDuplicateLines.Command")
    Application.CommandBars("PlotCalc").Controls(15).ToolTipText = localization.ActiveLng.Item("ConvertAllToCurves.Command")
    Application.CommandBars("PlotCalc").Controls(16).ToolTipText = localization.ActiveLng.Item("CommandBars.ToolTipText9")
End Sub
