Public Sub Load()
    If isUserDataLoad = True Then Exit Sub
    On Error GoTo 10
    Dim file As String
    file = VGCore.GMSManager.Projects("PlotCalc").FilePath & "plotcalc_userdata.bin"
    
    Dim version As Long
    Dim isLittleEndian As Boolean
    Dim buf As New buffer
    Dim crc As Long
    
    Dim bytes() As Byte
    ReDim bytes(fFileLen(file) - 1)

    Dim fileNumber As Long
    fileNumber = FreeFile
    Open file For Binary Access Read As fileNumber
        Get fileNumber, , bytes
    Close fileNumber
    buf.SetBytes bytes
    
    version = buf.bGet_Long
    isLittleEndian = buf.bGet_Boolean
    
    ToolBarVersion = buf.bGet_Long
    PerimetrFormData.Price = buf.bGet_Long
    PerimetrFormData.Discount = buf.bGet_Long
    PerimetrFormData.MoneyUnit = buf.bGet_String(10)
    PerimetrFormData.WindowPositionX = buf.bGet_Long
    PerimetrFormData.WindowPositionY = buf.bGet_Long
    SquareFormData.Price = buf.bGet_Long
    SquareFormData.WindowPositionX = buf.bGet_Long
    SquareFormData.WindowPositionY = buf.bGet_Long
    SquareFormData.RectArea = buf.bGet_Integer
    CodeOfActivation = buf.bGet_String(20)
    LangCode = buf.bGet_String(2)
    UnitsSystem = buf.bGet_Long
    AutoUpdate = buf.bGet_Boolean
    LastProfileSelectionNumber = buf.bGet_Long
    CountRaskroyProfiles = buf.bGet_Long
    
    If CountRaskroyProfiles > 0 Then
        ReDim RaskroyProfiles(CountRaskroyProfiles - 1)
        Dim q As Long
        For q = 0 To CountRaskroyProfiles - 1
            RaskroyProfiles(q).name = buf.bGet_String(100)
            RaskroyProfiles(q).MaterialWith = buf.bGet_Long
            RaskroyProfiles(q).MaterialHeight = buf.bGet_Long
            RaskroyProfiles(q).PlotterSheetWith = buf.bGet_Long
            RaskroyProfiles(q).PlotterSheetHeight = buf.bGet_Long
            
            RaskroyProfiles(q).LeftField = buf.bGet_Long
            RaskroyProfiles(q).RightField = buf.bGet_Long
            RaskroyProfiles(q).UpField = buf.bGet_Long
            RaskroyProfiles(q).DownField = buf.bGet_Long
            RaskroyProfiles(q).Otstup = buf.bGet_Long
            
            RaskroyProfiles(q).Vnutri_Figuri = buf.bGet_Boolean
            RaskroyProfiles(q).Vlogen = buf.bGet_Boolean
            RaskroyProfiles(q).GradusCount = buf.bGet_Byte
    
            RaskroyProfiles(q).DrawMaterial = buf.bGet_Boolean
            RaskroyProfiles(q).DrawFields = buf.bGet_Boolean
            RaskroyProfiles(q).DrawWorkObl = buf.bGet_Boolean
            RaskroyProfiles(q).DrawText = buf.bGet_Boolean
            RaskroyProfiles(q).PlotCalcLayer = buf.bGet_Boolean
            RaskroyProfiles(q).Group = buf.bGet_Boolean
            RaskroyProfiles(q).LockObj = buf.bGet_Boolean
        Next
    End If
    
    crc = buf.bGet_Long
    

    'Проверка подписи
    If CRC32(bytes) <> crc Then
        'Подпись не совпала
        'MsgBox "Podpis"
        Err.Raise 2
    End If

    'Проверка версии
    If version <> CurUserDataFileVersion Then
        'версия не совпала
        'MsgBox "Versia"
        Err.Raise 3
    End If
   
    isUserDataLoad = True
    
    localization.SelectLangByCode (LangCode)
    Exit Sub
10
    SetDeafault
End Sub
