Private Sub StartNesting()
    'Передать параметры и начать процесс размещения
    Dim To4nost As Long:                    To4nost = 20 'Количество полос
    Dim GradusCount2 As Long
    Dim Vlogen As Boolean

    SingMaterialWith = DelTochka(Text4) * 1000
    SingMaterialHeight = DelTochka(Text5) * 1000
    SingPlotterSheetWith = (DelTochka(Text8) - (DelTochka(Text13) + DelTochka(Text14))) * 1000
    SingPlotterSheetHeight = (DelTochka(Text7) - (DelTochka(Text12) + DelTochka(Text15))) * 1000
    SingLeftField = DelTochka(Text13) * 1000
    SingRightField = DelTochka(Text14) * 1000
    SingUpField = DelTochka(Text12) * 1000
    SingDownField = DelTochka(Text15) * 1000
    Otstup = DelTochka(Text18) / 2 * 1000

    Select Case Combo2.ListIndex
        Case 0
            GradusCount2 = 1
        Case 1
            GradusCount2 = 2
        Case 2
            GradusCount2 = 4
        Case 3
            GradusCount2 = 8
        Case Else
            GradusCount2 = 32  'Кол-во градусов
    End Select
    
    If Check1.value Then
        Vlogen = False 'Разрешить вложение
    Else
        Vlogen = True 'Запретить вложение
    End If
    
    If Check2.value Then
        Vnutri_Figuri = True 'Внутри самой Большой фигуры
        SingMaterialWith = AlfaWidth * 25.4 * 1000 + 10
        SingMaterialHeight = AlfaHeight * 25.4 * 1000 + 10
        SingPlotterSheetWith = AlfaWidth * 25.4 * 1000 + 10
        SingPlotterSheetHeight = AlfaHeight * 25.4 * 1000 + 10
        SingLeftField = 0
        SingRightField = 0
        SingUpField = 0
        SingDownField = 0
    Else
        Vnutri_Figuri = False 'Внутри самой Большой фигуры
    End If
        
    Dim buffer1 As New buffer
    With buffer1
        .InterPlatformNetConverter = True
        .size = 52
        .bPut_Long Commands.Place
        .bPut_Long GradusCount2
        .bPut_Long To4nost
        .bPut_Long Otstup
        .bPut_Long SingMaterialWith
        .bPut_Long SingMaterialHeight
        .bPut_Long SingPlotterSheetWith
        .bPut_Long SingPlotterSheetHeight
        .bPut_Long SingLeftField
        .bPut_Long SingRightField
        .bPut_Long SingUpField
        .bPut_Long SingDownField
        .bPut_Boolean Vnutri_Figuri
        .bPut_Boolean Vlogen
    End With
    BlocksSocket1.Send buffer1.bytes
End Sub
