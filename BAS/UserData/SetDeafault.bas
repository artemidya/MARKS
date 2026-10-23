Public Sub SetDeafault()
    LangCode = GetInfo
    localization.SelectLangByCode (LangCode)
    PerimetrFormData.MoneyUnit = GetMoneyUnit
    
    PerimetrFormData.Discount = -5000
    PerimetrFormData.Price = 0
    PerimetrFormData.WindowPositionX = 200
    PerimetrFormData.WindowPositionY = 200
    
    SquareFormData.Price = 0
    SquareFormData.WindowPositionX = 210
    SquareFormData.WindowPositionY = 210
    SquareFormData.RectArea = 0
    
    '-------------------
    LastProfileSelectionNumber = 0
    CountRaskroyProfiles = 7
    ReDim RaskroyProfiles(CountRaskroyProfiles - 1)
    
    With RaskroyProfiles(0)
       .name = localization.ActiveLng.Item("UserData.SetDeafault.RaskroyProfiles0")
       
       .MaterialHeight = 5000000
       .MaterialWith = 1000000
       .PlotterSheetHeight = 5000000
       .PlotterSheetWith = 1000000
       
       .LeftField = 25000
       .RightField = 25000
       .UpField = 25000
       .DownField = 25000

       .Otstup = 1000
       .GradusCount = 32
       
       .Vnutri_Figuri = False
    End With
    
    With RaskroyProfiles(1)
       .name = localization.ActiveLng.Item("UserData.SetDeafault.RaskroyProfiles1")
       
       .MaterialHeight = 5000000
       .MaterialWith = 1200000
       .PlotterSheetHeight = 580000
       .PlotterSheetWith = 1200000
       
       .LeftField = 25000
       .RightField = 25000
       .UpField = 25000
       .DownField = 25000

       .Otstup = 1000
       .GradusCount = 32
       
       .Vnutri_Figuri = False
    End With
    
    With RaskroyProfiles(2)
       .name = localization.ActiveLng.Item("UserData.SetDeafault.RaskroyProfiles2")
       
       .MaterialHeight = 5000000
       .MaterialWith = 1000000
       .PlotterSheetHeight = 5000000
       .PlotterSheetWith = 1000000
       
       .LeftField = 25000
       .RightField = 25000
       .UpField = 25000
       .DownField = 25000

       .Otstup = 5000
       .GradusCount = 4
       
       .Vnutri_Figuri = False
    End With
    
    With RaskroyProfiles(3)
       .name = localization.ActiveLng.Item("UserData.SetDeafault.RaskroyProfiles3")
       
       .MaterialHeight = 6000000
       .MaterialWith = 3000000
       .PlotterSheetHeight = 1500000
       .PlotterSheetWith = 1500000
       
       .LeftField = 50000
       .RightField = 50000
       .UpField = 50000
       .DownField = 50000

       .Otstup = 50000
       .GradusCount = 32
       
       .Vnutri_Figuri = False
    End With
    
    With RaskroyProfiles(4)
       .name = localization.ActiveLng.Item("UserData.SetDeafault.RaskroyProfiles4")
       
       .MaterialHeight = 5000000
       .MaterialWith = 1000000
       .PlotterSheetHeight = 580000
       .PlotterSheetWith = 1200000
       
       .LeftField = 25000
       .RightField = 25000
       .UpField = 25000
       .DownField = 25000

       .Otstup = 1000
       .GradusCount = 32
       
       .Vnutri_Figuri = False
    End With
    
    With RaskroyProfiles(5)
       .name = localization.ActiveLng.Item("UserData.SetDeafault.RaskroyProfiles5")
       
       .MaterialHeight = 5000000
       .MaterialWith = 1200000
       .PlotterSheetHeight = 5000000
       .PlotterSheetWith = 1200000
       
       .LeftField = 25000
       .RightField = 25000
       .UpField = 25000
       .DownField = 25000

       .Otstup = 1000
       .GradusCount = 32
       
       .Vnutri_Figuri = False
    End With
    
    With RaskroyProfiles(6)
       .name = localization.ActiveLng.Item("UserData.SetDeafault.RaskroyProfiles6")
       
       .MaterialHeight = 0
       .MaterialWith = 0
       .PlotterSheetHeight = 0
       .PlotterSheetWith = 0
       
       .LeftField = 0
       .RightField = 0
       .UpField = 0
       .DownField = 0

       .Otstup = 1000
       .GradusCount = 32
       
       .Vnutri_Figuri = True
    End With
    
    Dim q As Long
    For q = 0 To UBound(RaskroyProfiles)
        With RaskroyProfiles(q)
            .Vlogen = False
            .DrawFields = True
            .DrawMaterial = True
            .DrawText = True
            .DrawWorkObl = True
            .Group = True
            .LockObj = False
            .PlotCalcLayer = False
        End With
    Next
    CodeOfActivation = ""
    
    isUserDataLoad = True
    ToolBarVersion = 0 'сброс тулбара
    
    Dim RegData1 As New RegData
    RegData1.CropBMPoffset = 500
    RegData1.TrimOffset = 300
End Sub
