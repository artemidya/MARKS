Private Sub Razmestit_Reazultat(ByRef Work_Result As Work_Result, _
                         DrawMaterial As Boolean, _
                         DrawFields As Boolean, _
                         DrawWorkObl As Boolean, _
                         DrawText As Boolean, _
                         PlotCalcLayer As Boolean, _
                         Group As Boolean, _
                         LockObj As Boolean, _
                         OnPages As Boolean)
    '===================================================
    'Размещает фигуры в CorelDRAW
    '===================================================
    On Error GoTo 20
    VGCore.Optimization = True
    VGCore.ActiveDocument.BeginCommandGroup localization.ActiveLng.Item("RaskroyForm.BeginCommandGroup")
    Okrugl Work_Result 'Округление размеров
    Dim pi As Double: pi = 3.14159265
    Dim plusX As Double
    Dim plusY As Double
    Dim lastPX As Double, lastPY As Double
    Dim sheetX As Double, sheetY As Double
    Dim text2 As String
    Dim dx As Double, dy As Double
    Dim rightGo As Boolean
    Dim shiftX As Double, shiftY As Double
    Dim q As Long, m As Long, rl As Long, rs As Long
    
    Dim RazmetkaObject As VGCore.shape
    Dim RazmetkaObjects As New VGCore.ShapeRange
    Dim RaskroyObjects As New VGCore.ShapeRange
    
    Dim RazmLayer As VGCore.Layer
    Dim RazmLayerStr As String

    Dim CurPage As VGCore.page
    
    If OnPages Then
        With VGCore.ActiveDocument.ActivePage
            .SizeWidth = Work_Result.SingMaterialWith
            .SizeHeight = Work_Result.SingMaterialHeight
        
            If Not IsFreePlaceOnThePage(.LeftX, .BottomY, .SizeWidth, .SizeHeight) Then
                'Insert New Page
                ActiveDocument.InsertPagesEx 1, False, .index, Work_Result.SingMaterialWith, Work_Result.SingMaterialHeight
            End If
        End With
    Else
        lastPX = VGCore.ActiveSelection.PositionX + VGCore.ActiveSelection.SizeWidth + Work_Result.SingMaterialWith / 20
        lastPY = VGCore.ActiveSelection.PositionY - VGCore.ActiveSelection.SizeHeight
    End If
    
    If Work_Result.SingMaterialWith > Work_Result.SingMaterialHeight Then
        plusY = Work_Result.SingMaterialHeight + Work_Result.SingMaterialHeight / 20
        rightGo = False
    Else
        plusX = Work_Result.SingMaterialWith + Work_Result.SingMaterialWith / 20
        rightGo = True
    End If
    
    With Work_Result
        For m = 1 To .CountMaterials
            sheetY = 0
            With .Materials(m)
                If OnPages Then
                    .PositionX = VGCore.ActiveDocument.ActivePage.LeftX
                    .PositionY = VGCore.ActiveDocument.ActivePage.BottomY
                Else
                    If m > 1 Then
                        .PositionX = lastPX + plusX
                        .PositionY = lastPY + plusY
                    Else
                        .PositionX = lastPX
                        .PositionY = lastPY
                    End If
                    'Проверить свободна ли позиция
                    FindPostion .PositionX, .PositionY, Work_Result.SingMaterialWith, Work_Result.SingMaterialHeight, rightGo
                    lastPX = .PositionX
                    lastPY = .PositionY
                End If
                
                If DrawMaterial Then
                    Set RazmetkaObject = VGCore.ActiveLayer.CreateRectangle(.PositionX, .PositionY, Work_Result.SingMaterialWith + .PositionX, Work_Result.SingMaterialHeight + .PositionY) 'Это весь материал
                    RazmetkaObject.Outline.SetProperties Color:=CreateRGBColor(102, 102, 204)
                    RazmetkaObjects.Add RazmetkaObject
                End If
                For rl = 1 To .CountRLines
                    sheetX = 0
                    With .RLines(rl)
                        'текст==========================================================================
                        text2 = "/"
                        If DrawText Then
                            If .CountRSheets = 1 Then
                                If .RSheets(1).CountRFigures > 0 Then
                                    Set RazmetkaObject = VGCore.ActiveLayer.CreateArtisticText(0, 0, Int((.Height + Work_Result.SingDownField + Work_Result.SingUpField) * 25.4 + 0.5))
                                    RazmetkaObject.Rotate 90
                                    RazmetkaObject.Stretch Work_Result.SingMaterialHeight / 20
                                    RazmetkaObject.PositionY = sheetY + Work_Result.Materials(m).PositionY + (.Height + Work_Result.SingDownField + Work_Result.SingUpField) / 2 + RazmetkaObject.SizeHeight / 2
                                    RazmetkaObject.PositionX = Work_Result.Materials(m).PositionX - RazmetkaObject.SizeWidth - RazmetkaObject.SizeWidth / 2
                                    RazmetkaObjects.Add RazmetkaObject
                                End If
                            Else
                                If .RSheets(1).CountRFigures > 0 Then
                                    Set RazmetkaObject = VGCore.ActiveLayer.CreateArtisticText(0, 0, Int((.Height + Work_Result.SingDownField + Work_Result.SingUpField) * 25.4 + 0.5))
                                    RazmetkaObject.Stretch Work_Result.SingMaterialHeight / 20
                                    RazmetkaObject.Rotate 270
                                    RazmetkaObject.PositionY = sheetY + Work_Result.Materials(m).PositionY + (.Height + Work_Result.SingDownField + Work_Result.SingUpField) / 2 + RazmetkaObject.SizeHeight / 2
                                    RazmetkaObject.PositionX = Work_Result.Materials(m).PositionX + Work_Result.SingMaterialWith + RazmetkaObject.SizeWidth / 2
                                    dx = RazmetkaObject.PositionX + RazmetkaObject.SizeWidth
                                    dy = RazmetkaObject.PositionY - RazmetkaObject.SizeHeight
                                    RazmetkaObjects.Add RazmetkaObject
                                End If
                            End If
                        End If '========================================================================
                        For rs = 1 To .CountRSheets
                            With .RSheets(rs)
                                If .CountRFigures > 0 Then
                                    text2 = text2 & Int((Work_Result.SingLeftField + .Width + Work_Result.SingRightField) * 25.4 + 0.5) & "/"
                                    If DrawFields Then

                                        Set RazmetkaObject = VGCore.ActiveLayer.CreateRectangle(Work_Result.Materials(m).PositionX + Work_Result.SingLeftField + sheetX, _
                                                                                    Work_Result.Materials(m).PositionY + Work_Result.SingDownField + sheetY, _
                                                                                   .Width + Work_Result.Materials(m).PositionX + Work_Result.SingLeftField + sheetX, _
                                                                                    Work_Result.Materials(m).RLines(rl).Height + Work_Result.Materials(m).PositionY + Work_Result.SingDownField + sheetY) 'Это весь материал
                                        RazmetkaObject.Outline.SetProperties Color:=CreateRGBColor(255, 102, 0)
                                        RazmetkaObject.Outline.SetProperties 0.059055
                                        RazmetkaObjects.Add RazmetkaObject
                                    End If
                                    If DrawWorkObl Then
                                        Set RazmetkaObject = VGCore.ActiveLayer.CreateRectangle(Work_Result.Materials(m).PositionX + sheetX, _
                                                                                        Work_Result.Materials(m).PositionY + sheetY, _
                                                                                       .Width + Work_Result.Materials(m).PositionX + Work_Result.SingLeftField + Work_Result.SingRightField + sheetX, _
                                                                                        Work_Result.Materials(m).RLines(rl).Height + Work_Result.Materials(m).PositionY + Work_Result.SingDownField + Work_Result.SingUpField + sheetY) 'Это весь материал
                                        RazmetkaObject.Outline.SetProperties Color:=CreateRGBColor(255, 102, 0)
                                        RazmetkaObjects.Add RazmetkaObject
                                    End If
                                End If
                                
                                For q = 1 To .CountRFigures
                                    With .RFigures(q)
                                        Rerotator shiftX, shiftY, RerotateCrv(.Q_figure), Du(.Q_figure), (-.gradus / (2 * pi / 360))
                                        Du(.Q_figure).PositionX = .x - Du(.Q_figure).SizeWidth + Work_Result.SingLeftField + Work_Result.Materials(m).PositionX + sheetX - 2 * Work_Result.Otstup - shiftX
                                        Du(.Q_figure).PositionY = .y + Work_Result.SingDownField + Work_Result.Materials(m).PositionY + sheetY - 2 * Work_Result.Otstup - shiftY
                                        RaskroyObjects.AddRange Du(.Q_figure)
                                    End With
                                Next
                                
                                
                                sheetX = sheetX + Work_Result.SingLeftField + .Width + Work_Result.SingRightField
                            End With
                        Next
                        sheetY = sheetY + Work_Result.SingDownField + .Height + Work_Result.SingUpField
                        'текст горизонтальный
                        If DrawText And .CountRSheets > 1 And .RSheets(1).CountRFigures > 0 Then
                            Set RazmetkaObject = ActiveLayer.CreateArtisticText(dx, dy, text2)
                            RazmetkaObject.Stretch Work_Result.SingMaterialHeight / 20
                            RazmetkaObjects.Add RazmetkaObject
                        End If
                        
                        
                        
                    End With
                Next
            End With
            'RazmLayerStr
            If OnPages Then
                If m = 1 Then
                    Set CurPage = VGCore.ActivePage
                    
                    If PlotCalcLayer Then
                        VGCore.ActivePage.CreateLayer "PlotCalc"
                        RazmLayerStr = "PlotCalc"
                        
                    Else
                        RazmLayerStr = VGCore.ActiveLayer.name
                    End If
                Else ' m > 1
                    Set CurPage = ActiveDocument.InsertPagesEx(1, False, VGCore.ActivePage.index, .SingMaterialWith, .SingMaterialHeight)
                    
                End If
                
                
                RazmetkaObjects.MoveToLayer CurPage.Layers(RazmLayerStr)
                For q = 1 To RaskroyObjects.count
                    With RaskroyObjects(q)
                        .MoveToLayer CurPage.Layers(.Layer.name)
                    End With
                Next
                
                VGCore.ActiveDocument.ClearSelection
                If RazmetkaObjects.count > 0 Then
                    RazmetkaObjects.AddToSelection
                    If Group Then
                        VGCore.ActiveSelection.Group
                    End If
                    If LockObj Then
                        VGCore.ActiveSelection.Locked = True
                    End If
                End If
                
                RazmetkaObjects.RemoveAll
                RaskroyObjects.RemoveAll
            End If
        Next
    End With
    
    VGCore.ActiveDocument.ClearSelection
    
    If Not OnPages Then
        If RazmetkaObjects.count > 0 Then
            'Layer
            
            If PlotCalcLayer Then
                Set RazmLayer = VGCore.ActivePage.CreateLayer("PlotCalc")
                RazmetkaObjects.MoveToLayer RazmLayer
            End If
            
            RazmetkaObjects.AddToSelection
            If Group Then
                VGCore.ActiveSelection.Group
            End If
            If LockObj Then
                VGCore.ActiveSelection.Locked = True
            End If
        End If
        
        'Добаить в выделение ВСЁ
        VGCore.ActiveDocument.ClearSelection
        RazmetkaObjects.AddToSelection
        RaskroyObjects.AddToSelection
    End If


20
    VGCore.Optimization = False
    VGCore.ActiveDocument.EndCommandGroup
    VGCore.Refresh
End Sub
