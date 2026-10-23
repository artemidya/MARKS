Private Sub DrawWarkResult(ByRef Work_Result As Work_Result)
    DrawWarkResultProcessing = True
    On Error Resume Next
    Dim ks As Double
    Dim dx As Double, dy As Double
    Dim oblastH As Double, oblastW As Double
    Dim minusH As Double, minusW As Double
    Dim horisontal As Boolean 'Горихонталь или вертикаль
    Dim plusY As Double, plusX As Double
    minusH = 0: minusW = 0
    plusX = 0: plusY = 0
    If Work_Result.SingMaterialHeight < Work_Result.SingMaterialWith Then
        oblastH = Work_Result.SingMaterialHeight * Work_Result.CountMaterials
        oblastW = Work_Result.SingMaterialWith
        minusH = 5 * Work_Result.CountMaterials - 5
        horisontal = False
    Else
        oblastH = Work_Result.SingMaterialHeight
        oblastW = Work_Result.SingMaterialWith * Work_Result.CountMaterials
        minusW = 5 * Work_Result.CountMaterials - 5
        horisontal = True
    End If
    Cls
    If oblastW / WindowWidth > oblastH / WindowHeigth Then
        ks = (WindowWidth - 15 - minusW) / oblastW
    Else
        ks = (WindowHeigth - 20 - minusH) / oblastH
    End If
    dx = 6
    dy = 8
    Dim q As Long
    'Теперь нарисовать фигуры ===================================
    Dim q2 As Long
    Dim q3 As Long
    Dim q4 As Long
    Dim xx1 As Double, w As Long
    Dim yy1 As Double, h As Long
    Dim hh As Long, hh2 As Long
    w = Work_Result.SingMaterialWith * ks
    h = Work_Result.SingMaterialHeight * ks
    plusX = 0
    plusY = 0
    For q = 1 To Work_Result.CountMaterials
        With Work_Result.Materials(q)
            If horisontal Then
                plusX = (q - 1) * Work_Result.SingMaterialWith + (q - 1) * 5 / ks
                plusY = Work_Result.SingDownField
            Else
                plusX = 0
                plusY = (q - 1) * Work_Result.SingMaterialHeight + Work_Result.SingDownField + (q - 1) * 5 / ks
            End If
            xx1 = dx + plusX * ks
            yy1 = WindowHeigth - dy - plusY * ks + Work_Result.SingDownField * ks - Work_Result.SingMaterialHeight * ks
            If q = 1 And Work_Result.Vnutri_Figuri = True Then
                DrawFigList ks
            Else
                FillRect ShadowBrush, xx1 + 2, yy1 + 2, w, h
                FillRect PaperBrush, xx1, yy1, w, h
                DrawRect LinePen, xx1, yy1, w, h
            End If
            For q2 = 1 To .CountRLines
                With .RLines(q2)
                    If horisontal Then
                        plusX = (q - 1) * Work_Result.SingMaterialWith + Work_Result.SingLeftField + (q - 1) * 5 / ks
                    Else
                        plusX = Work_Result.SingLeftField
                    End If
                    hh = (Work_Result.SingUpField + .Height + Work_Result.SingDownField) * ks
                    hh2 = .Height * ks
                    For q3 = 1 To .CountRSheets
                        With .RSheets(q3)
                            For q4 = 1 To .CountRFigures
                                With .RFigures(q4)
                                    DrawFigure RCurves(.Q_figure), .x + plusX - Work_Result.Otstup * 2, .y + plusY - Work_Result.Otstup * 2, .gradus, ks
                                End With
                                DoEvents
                                If Not DrawWarkResultProcessing Then Exit Sub
                            Next
                            'Нарисовать границы области
                            If .CountRFigures > 0 And Work_Result.Vnutri_Figuri = False Then
                                xx1 = dx + plusX * ks - Work_Result.SingLeftField * ks
                                yy1 = WindowHeigth - dy - plusY * ks - hh + Work_Result.SingDownField * ks
                                DrawRect FieldsPen, xx1 + Work_Result.SingLeftField * ks, yy1 + Work_Result.SingUpField * ks, .Width * ks, hh2
                                DrawRect PagesPen, xx1, yy1, (.Width + Work_Result.SingLeftField + Work_Result.SingRightField) * ks, hh
                            End If
                            plusX = plusX + Work_Result.SingRightField + .Width + Work_Result.SingLeftField
                        End With
                    Next
                    plusY = plusY + Work_Result.SingDownField + .Height + Work_Result.SingUpField
                End With
            Next
        End With
    Next
    If Work_Result.Errors = True Then
        Dim str As String
        str = localization.ActiveLng.Item("RaskroyForm.NOTE_SomeCharactersNotPossible")
        TextOut HDCForm, shiftX + 10, shiftY + WindowHeigth - 30, str, Len(str)
    End If
    DrawWarkResultProcessing = False
End Sub
