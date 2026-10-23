Private Sub Okrugl(ByRef Work_Result As Work_Result)
    '===========================================
    'Округляет результаты (размеры страниц)
    '===========================================
    Dim q As Long, q2 As Long, q3 As Long
    With Work_Result
        If .Materials(1).RLines(1).CountRSheets = 1 Then
            For q = 1 To .CountMaterials
                With .Materials(q)
                    For q2 = 1 To .CountRLines
                        With .RLines(q2)
                            .RSheets(1).Width = Work_Result.SingMaterialWith - Work_Result.SingLeftField - Work_Result.SingRightField
                        End With
                    Next
                End With
            Next
        End If
    End With
    'теперь округление
    Dim ostatok  As Double
    Dim d As Double, d2 As Double, d3 As Double
    With Work_Result
            For q = 1 To .CountMaterials
                With .Materials(q)
                    For q2 = 1 To .CountRLines
                        With .RLines(q2)
                            'Находим остаток
                            ostatok = Work_Result.SingMaterialWith
                            For q3 = 1 To .CountRSheets
                                With .RSheets(q3)
                                    ostatok = ostatok - .Width - Work_Result.SingLeftField - Work_Result.SingRightField
                                End With
                            Next
                            'Тепреь распределим остаток на все страницы так чтобы округлить их
                            For q3 = 1 To .CountRSheets
                                With .RSheets(q3)
                                    d = ((1 - ((.Width * 25.4) / 10 - Int((.Width * 25.4) / 10))) * 10) / 25.4  'то сколько нужно прибывить (кратно 10)
                                    d2 = ((1 - ((.Width * 25.4) / 5 - Int((.Width * 25.4) / 5))) * 5) / 25.4    'то сколько нужно прибывить (кратно 5)
                                    d3 = ((1 - ((.Width * 25.4) / 1 - Int((.Width * 25.4) / 1))) * 1) / 25.4    'то сколько нужно прибывить (кратно 1)
                                    If ostatok > d And (.Width + d <= Work_Result.SingPlotterSheetWith) Then
                                        ostatok = ostatok - d
                                        .Width = .Width + d
                                    ElseIf ostatok > d2 And (.Width + d2 <= Work_Result.SingPlotterSheetWith) Then
                                         ostatok = ostatok - d2
                                        .Width = .Width + d2
                                    ElseIf ostatok > d3 And (.Width + d3 <= Work_Result.SingPlotterSheetWith) Then
                                         ostatok = ostatok - d3
                                        .Width = .Width + d3
                                    End If
                                End With
                            Next
                        End With
                    Next
                    'теперь вертикальное округление
                    ostatok = Work_Result.SingMaterialHeight
                    'находим остаток
                    For q2 = 1 To .CountRLines
                        With .RLines(q2)
                            ostatok = ostatok - .Height - Work_Result.SingUpField - Work_Result.SingDownField
                        End With
                    Next
                    
                    For q2 = 1 To .CountRLines
                        With .RLines(q2)
                            d = ((1 - ((.Height * 25.4) / 10 - Int((.Height * 25.4) / 10))) * 10) / 25.4    'то сколько нужно прибывить (кратно 10)
                            d2 = ((1 - ((.Height * 25.4) / 5 - Int((.Height * 25.4) / 5))) * 5) / 25.4      'то сколько нужно прибывить (кратно 5)
                            d3 = ((1 - ((.Height * 25.4) / 1 - Int((.Height * 25.4) / 1))) * 1) / 25.4      'то сколько нужно прибывить (кратно 1)
                            If ostatok > d And (.Height + d <= Work_Result.SingPlotterSheetHeight) Then
                                ostatok = ostatok - d
                                .Height = .Height + d
                            ElseIf ostatok > d2 And (.Height + d2 <= Work_Result.SingPlotterSheetHeight) Then
                                ostatok = ostatok - d2
                                .Height = .Height + d2
                            ElseIf ostatok > d3 And (.Height + d3 <= Work_Result.SingPlotterSheetHeight) Then
                                ostatok = ostatok - d3
                                .Height = .Height + d3
                            End If
                        End With
                    Next
                End With
            Next
    End With
End Sub
