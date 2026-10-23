Public Sub Save()
    On Error GoTo 10
    If isUserDataLoad = False Then Exit Sub
    Dim file As String: file = VGCore.GMSManager.Projects("PlotCalc").FilePath & "plotcalc_userdata.bin"
    Dim crc As Long
    Dim buf As New buffer
    buf.size = 4 + _
               2 + _
               LenB(ToolBarVersion) + _
               LenB(PerimetrFormData.Price) + _
               LenB(PerimetrFormData.Discount) + _
               LenB(PerimetrFormData.MoneyUnit) + _
               LenB(PerimetrFormData.WindowPositionX) + _
               LenB(PerimetrFormData.WindowPositionY) + _
               LenB(SquareFormData.Price) + _
               LenB(SquareFormData.WindowPositionX) + _
               LenB(SquareFormData.WindowPositionY) + _
               LenB(SquareFormData.RectArea) + _
               LenB(CodeOfActivation) + _
               LenB(LangCode) + _
               LenB(UnitsSystem) + _
               LenB(AutoUpdate) + _
               LenB(LastProfileSelectionNumber) + _
               LenB(CountRaskroyProfiles) + _
               LenB(crc) + _
               255 * CountRaskroyProfiles

    buf.bPut_Long CurUserDataFileVersion
    buf.bPut_Boolean isLittleEndianf()
    buf.bPut_Long ToolBarVersion
    buf.bPut_Long PerimetrFormData.Price
    buf.bPut_Long PerimetrFormData.Discount
    buf.bPut_String PerimetrFormData.MoneyUnit
    buf.bPut_Long PerimetrFormData.WindowPositionX
    buf.bPut_Long PerimetrFormData.WindowPositionY
    buf.bPut_Long SquareFormData.Price
    buf.bPut_Long SquareFormData.WindowPositionX
    buf.bPut_Long SquareFormData.WindowPositionY
    buf.bPut_Integer SquareFormData.RectArea
    buf.bPut_String CodeOfActivation
    buf.bPut_String LangCode
    buf.bPut_Long UnitsSystem
    buf.bPut_Boolean AutoUpdate
    buf.bPut_Long LastProfileSelectionNumber
    buf.bPut_Long CountRaskroyProfiles
    
    Dim q As Long
    For q = 0 To CountRaskroyProfiles - 1
        buf.bPut_String RaskroyProfiles(q).name '200
        buf.bPut_Long RaskroyProfiles(q).MaterialWith   '4
        buf.bPut_Long RaskroyProfiles(q).MaterialHeight '4
        buf.bPut_Long RaskroyProfiles(q).PlotterSheetWith '4
        buf.bPut_Long RaskroyProfiles(q).PlotterSheetHeight '4
        
        buf.bPut_Long RaskroyProfiles(q).LeftField '4
        buf.bPut_Long RaskroyProfiles(q).RightField '4
        buf.bPut_Long RaskroyProfiles(q).UpField '4
        buf.bPut_Long RaskroyProfiles(q).DownField '4
        buf.bPut_Long RaskroyProfiles(q).Otstup '4
        
        buf.bPut_Boolean RaskroyProfiles(q).Vnutri_Figuri '2
        buf.bPut_Boolean RaskroyProfiles(q).Vlogen '2
        buf.bPut_Byte RaskroyProfiles(q).GradusCount '1

        buf.bPut_Boolean RaskroyProfiles(q).DrawMaterial '2
        buf.bPut_Boolean RaskroyProfiles(q).DrawFields '2
        buf.bPut_Boolean RaskroyProfiles(q).DrawWorkObl '2
        buf.bPut_Boolean RaskroyProfiles(q).DrawText '2
        buf.bPut_Boolean RaskroyProfiles(q).PlotCalcLayer '2
        buf.bPut_Boolean RaskroyProfiles(q).Group '2
        buf.bPut_Boolean RaskroyProfiles(q).LockObj '2
    Next

    crc = CRC32(buf.bytes)
    buf.bPut_Long crc
    
    If fFileLen(file) > buf.size Then KillFile file
    
    Dim fileNumber As Long
    fileNumber = FreeFile
    Open file For Binary Access Write As fileNumber
        Put fileNumber, , buf.bytes
    Close fileNumber
    
    Exit Sub
10
    MsgBox "(ошибка 001) Не возможно сохранить данные пользователя в файл " & file, vbCritical, "PlotCalc"
End Sub
