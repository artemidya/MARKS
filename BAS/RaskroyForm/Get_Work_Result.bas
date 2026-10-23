Private Function Get_Work_Result(ByRef buffer1 As buffer) As Work_Result
    On Error GoTo 50
    Dim q As Long, q2 As Long, q3 As Long, q4 As Long
    Dim errFlag As Long
    With Get_Work_Result
        .CountMaterials = buffer1.bGet_Long
        If buffer1.bGet_Long = 1 Then .Errors = True
        .maximumY = CDbl(buffer1.bGet_Long) / 25.4 / 1000
        ReDim .Materials(1 To .CountMaterials)
        For q = 1 To .CountMaterials
            With .Materials(q)
                .CountRLines = buffer1.bGet_Long
                ReDim .RLines(1 To .CountRLines)
                For q2 = 1 To .CountRLines
                    With .RLines(q2)
                        .CountRSheets = buffer1.bGet_Long
                        .Height = CDbl(buffer1.bGet_Long) / 25.4 / 1000
                        ReDim .RSheets(1 To .CountRSheets)
                        For q3 = 1 To .CountRSheets
                            With .RSheets(q3)
                                .CountRFigures = buffer1.bGet_Long
                                .Width = CDbl(buffer1.bGet_Long) / 25.4 / 1000
                                If .CountRFigures > 0 Then
                                ReDim .RFigures(1 To .CountRFigures)
                                    For q4 = 1 To .CountRFigures
                                        With .RFigures(q4)
                                            .gradus = CDbl(buffer1.bGet_Long) / 100000000
                                            .Q_figure = buffer1.bGet_Long
                                            .x = CDbl(buffer1.bGet_Long) / 25.4 / 1000
                                            .y = CDbl(buffer1.bGet_Long) / 25.4 / 1000
                                        End With
                                    Next
                                End If
                            End With
                        Next
                    End With
                Next
            End With
        Next
    End With
50
    With Get_Work_Result
        .SingMaterialWith = (SingMaterialWith / 1000) / 25.4
        .SingMaterialHeight = (SingMaterialHeight / 1000) / 25.4
        .SingPlotterSheetWith = (SingPlotterSheetWith / 1000) / 25.4
        .SingPlotterSheetHeight = (SingPlotterSheetHeight / 1000) / 25.4
        .SingLeftField = (SingLeftField / 1000) / 25.4
        .SingRightField = (SingRightField / 1000) / 25.4
        .SingUpField = (SingUpField / 1000) / 25.4
        .SingDownField = (SingDownField / 1000) / 25.4
        .Otstup = (Otstup / 1000) / 25.4
        .Vnutri_Figuri = Vnutri_Figuri
    End With
End Function
