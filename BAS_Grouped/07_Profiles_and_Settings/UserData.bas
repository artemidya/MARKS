' =============================================================================
' Group: 07_Profiles_and_Settings
' Source module: UserData.bas
' Procedures: 11
' =============================================================================

' -----------------------------------------------------------------------------
' Save (Sub), source lines 143-230
' -----------------------------------------------------------------------------
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

' -----------------------------------------------------------------------------
' Load (Sub), source lines 231-325
' -----------------------------------------------------------------------------
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

' -----------------------------------------------------------------------------
' SetDeafault (Sub), source lines 326-500
' -----------------------------------------------------------------------------
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

' -----------------------------------------------------------------------------
' GetMoneyUnit (Function), source lines 501-512
' -----------------------------------------------------------------------------
Private Function GetMoneyUnit() As String
    On Error GoTo 10
    GetMoneyUnit = "$"
    Dim buffer As String, ret As Long
    buffer = String(256, 0)
    ret = GetLocaleInfo(1024, &H14, buffer, Len(buffer))
    If ret > 0 Then
        GetMoneyUnit = left(buffer, ret - 1)
    End If
    Exit Function
10:
End Function

' -----------------------------------------------------------------------------
' GetInfo (Function), source lines 513-525
' -----------------------------------------------------------------------------
Private Function GetInfo() As String
    On Error GoTo 10
    Dim buffer As String, ret As Long
    buffer = String(256, 0)
    ret = GetLocaleInfo(1024, &H5C, buffer, Len(buffer))
    If ret > 1 Then
        GetInfo = left(buffer, 2)
    Else
        GetInfo = ""
    End If
    Exit Function
10:
End Function

' -----------------------------------------------------------------------------
' SortProfiles (Sub), source lines 526-546
' -----------------------------------------------------------------------------
Public Sub SortProfiles()
    Dim i As Long
    Dim fSwitched As Boolean
    Dim raskroyProfileTemp As RaskroyProfile
        Do
            fSwitched = False
            For i = 0 To CountRaskroyProfiles - 2
                If addZerroForDigits(UCase(RaskroyProfiles(i).name)) > addZerroForDigits(UCase(RaskroyProfiles(i + 1).name)) Then
                    fSwitched = True
                    raskroyProfileTemp = RaskroyProfiles(i)
                    RaskroyProfiles(i) = RaskroyProfiles(i + 1)
                    RaskroyProfiles(i + 1) = raskroyProfileTemp
                    If i = LastProfileSelectionNumber Then
                        LastProfileSelectionNumber = i + 1
                    ElseIf i + 1 = LastProfileSelectionNumber Then
                        LastProfileSelectionNumber = i
                    End If
                End If
            Next
        Loop While fSwitched
End Sub

' -----------------------------------------------------------------------------
' addZerroForDigits (Function), source lines 547-607
' -----------------------------------------------------------------------------
Private Function addZerroForDigits(ByVal txt As String, Optional ByVal maxDigits As Long = 50) As String
    txt = Replace(txt, ",", ".")
    Dim q As Long
    Dim code As Integer
    Dim lastCode As Integer
    Dim isStartingDigit As Boolean
    Dim startPositionDigits() As Long
    Dim endPositionDigits() As Long
    Dim digitsCount As Long
    
    Dim toRight() As Boolean ' (если дробь, то добавить нули справа)
    
    For q = 1 To Len(txt)
        code = Asc(Mid(txt, q, 1))
        If (code >= 48 And code <= 57) Then '..  Or code = 46
            If isStartingDigit = False Then
                isStartingDigit = True
                digitsCount = digitsCount + 1
                ReDim Preserve startPositionDigits(digitsCount - 1)
                ReDim Preserve endPositionDigits(digitsCount - 1)
                ReDim Preserve toRight(digitsCount - 1)
                startPositionDigits(digitsCount - 1) = q
                If lastCode = 46 Then toRight(digitsCount - 1) = True
            End If
        Else
            If isStartingDigit = True Then
                isStartingDigit = False
                endPositionDigits(digitsCount - 1) = q
            End If
        End If
        lastCode = code
    Next
    
    If isStartingDigit = True Then endPositionDigits(digitsCount - 1) = q
    
    Dim s As Long: s = 1
    Dim l As Long
    Dim l2 As Long
    Dim newDigits As String
    Dim q2 As Long

    For q = 0 To digitsCount - 1
        l = startPositionDigits(q) - s
        addZerroForDigits = addZerroForDigits & Mid(txt, s, l)
        l2 = endPositionDigits(q) - startPositionDigits(q)
        s = s + l + l2
        
        newDigits = ""
        For q2 = 1 To maxDigits - l2
            newDigits = newDigits & "0"
        Next
        If toRight(q) = True Then
            newDigits = Mid(txt, startPositionDigits(q), l2) & newDigits
        Else
            newDigits = newDigits & Mid(txt, startPositionDigits(q), l2)
        End If
        addZerroForDigits = addZerroForDigits & newDigits
    Next
    l = Len(txt) - s + 1
    addZerroForDigits = addZerroForDigits & Mid(txt, s, l)
End Function

' -----------------------------------------------------------------------------
' CRC32 (Function), source lines 608-617
' -----------------------------------------------------------------------------
Private Function CRC32(ByRef bytes() As Byte) As Long
    'Вычисляет контрольную сумму байтов (файла или пакета)
    'Работает до 8Мб
    On Error Resume Next
    Dim checkSum As Long, a As Long
    For a = 0 To UBound(bytes) - 4
       checkSum = checkSum + bytes(a)
    Next
    CRC32 = checkSum
End Function

' -----------------------------------------------------------------------------
' isLittleEndianf (Function), source lines 618-624
' -----------------------------------------------------------------------------
Private Function isLittleEndianf() As Boolean
    'тестирует формат записи int32 littleEndian или BigEndian в памяти
    Dim numberForTest As Long, bytes(3) As Byte
    numberForTest = 38778244
    CopyMemory bytes(0), numberForTest, 4
    If bytes(0) = 132 Then isLittleEndianf = True
End Function

' -----------------------------------------------------------------------------
' KillFile (Sub), source lines 625-629
' -----------------------------------------------------------------------------
Private Sub KillFile(fileName As String)
    On Error GoTo 10
    Kill fileName
10
End Sub

' -----------------------------------------------------------------------------
' fFileLen (Function), source lines 630-635
' -----------------------------------------------------------------------------
Private Function fFileLen(file As String) As Long
    On Error GoTo 10
    fFileLen = FileLen(file)
    Exit Function
10
End Function
