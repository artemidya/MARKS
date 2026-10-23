Private Sub SaveProfile(ByVal i As Long)
    With UserData.RaskroyProfiles(i)
        .name = Combo1_Text
    
        .MaterialWith = ConvertTextToLong(Text4.text)
        .MaterialHeight = ConvertTextToLong(Text5.text)
        .PlotterSheetWith = ConvertTextToLong(Text8.text)
        .PlotterSheetHeight = ConvertTextToLong(Text7.text)
        .UpField = ConvertTextToLong(Text12.text)
        .LeftField = ConvertTextToLong(Text13.text)
        .RightField = ConvertTextToLong(Text14.text)
        .DownField = ConvertTextToLong(Text15.text)
        .Otstup = ConvertTextToLong(Text18.text)
        
        .Vnutri_Figuri = Check2.value
        .Vlogen = Check1.value
        
        If Combo2.ListIndex = 0 Then
            .GradusCount = 1
        ElseIf Combo2.ListIndex = 1 Then
            .GradusCount = 2
        ElseIf Combo2.ListIndex = 2 Then
            .GradusCount = 4
        ElseIf Combo2.ListIndex = 3 Then
            .GradusCount = 8
        Else
            .GradusCount = 32
        End If
        
        .DrawMaterial = Check4.value
        .DrawWorkObl = Check3.value
        .DrawFields = Check5.value
        .DrawText = Check6.value
        .PlotCalcLayer = Check7.value
        .Group = Check8.value
        .LockObj = Check9.value
    End With
    Combo1.List(i) = Combo1_Text
    UserData.LastProfileSelectionNumber = Combo1LastSelectListIndex
    UserData.Save
End Sub
