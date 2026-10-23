Private Sub CreateToolBarForCorelDrawX8()
    On Error GoTo 10
    
    'Proverit versiu, esli ne sovpadaet to udalit PlotCalc ToolBar
    UserData.Load
    
    If UserData.ToolBarVersion <> CurToolBarVersion Then DeleteFilesAndToolBar
    '=============================================================================================================
    If IsPlotCalcToolBarExists = False Then
        MsgBox "Creating panel for PlotCalc", vbInformation
        Application.CommandBars.Add "PlotCalc"
        Init
        Dim q As Long
        For q = 0 To UBound(Subs)
            Application.CommandBars("PlotCalc").Controls.AddCustomButton UICategories.cdrCmdCategoryMacros, Subs(q)
        Next
        'to save cur toolbar version number
        UserData.ToolBarVersion = CurToolBarVersion
        UserData.Save
    End If
    If Application.CommandBars("PlotCalc").Visible = False Then
        Application.CommandBars("PlotCalc").Visible = True
    End If
    SetToolTipText
    SetIconsPictures
    Exit Sub
10
    MsgBox "(ошибка 002) не удается создать панель инструментов PlotCalc", vbCritical
End Sub
