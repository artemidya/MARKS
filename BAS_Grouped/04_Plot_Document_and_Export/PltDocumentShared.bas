' =============================================================================
' Group: 04_Plot_Document_and_Export
' Source module: PltDocumentShared.bas
' Procedures: 19
' =============================================================================

' -----------------------------------------------------------------------------
' GetPlotterFromString (Function), source lines 34-63
' -----------------------------------------------------------------------------
Public Function GetPlotterFromString(ByRef PlotterString As String, ByVal PlotterName As String) As Plotter2DModel

    Dim Plotter2DModelCacheCls As Plotter2DModelCache
    PlotterName = PlotterName & CRC32(PlotterString)

    Set Plotter2DModelCacheCls = GetItemFromPlottersCacheCol(PlotterName)
    
    If Plotter2DModelCacheCls Is Nothing Then
        Dim Convert As New ConverterStringLToPlotCalcCrv
        With GetPlotterFromString
            Set .Curves = Convert.GetPlotCalcCrv(GetTag("PlotterImg", PlotterString))
            Set .CarriageCurves = Convert.GetPlotCalcCrv(GetTag("CarriageImg", PlotterString))
            Dim boundingBox As RECTL
            boundingBox = CalculatePlotterBoundingBox(.Curves)
            .top = boundingBox.top
            .bottom = boundingBox.bottom
            .left = boundingBox.left
            .right = boundingBox.right
            .Height = .top - .bottom
            .Width = .right - .left
        End With
        Set Plotter2DModelCacheCls = New Plotter2DModelCache
        Plotter2DModelCacheCls.Plotter2DModel = GetPlotterFromString
        PlottersCacheCol.Add Plotter2DModelCacheCls, PlotterName

    Else
        GetPlotterFromString = Plotter2DModelCacheCls.Plotter2DModel
    End If

End Function

' -----------------------------------------------------------------------------
' CRC32 (Function), source lines 64-73
' -----------------------------------------------------------------------------
Private Function CRC32(ByRef txtForSignature As String) As String
    'Вычисляет контрольную сумму пакета
    'Работает до 8Мб
    On Error Resume Next
    Dim checkSum As Long, a As Long
    For a = 1 To Len(txtForSignature)
       checkSum = checkSum + Asc(Mid(txtForSignature, a, 1))
    Next
    CRC32 = checkSum
End Function

' -----------------------------------------------------------------------------
' ResetCache (Sub), source lines 74-77
' -----------------------------------------------------------------------------
Public Sub ResetCache()
    'Need reset cache when updating plotter, for updating plotter image
    Set PlottersCacheCol = New Collection
End Sub

' -----------------------------------------------------------------------------
' GetTag (Function), source lines 78-87
' -----------------------------------------------------------------------------
Private Function GetTag(TagName As String, ByRef FromText As String) As String
      Dim s As String, e As String
      s = InStr(1, FromText, "<" & TagName & ">", vbTextCompare)
      If s <> 0 Then
            e = InStr(s + Len(TagName) + 2, FromText, "</" & TagName & ">", vbTextCompare)
            If e <> 0 Then
                  GetTag = Mid(FromText, s + Len(TagName) + 2, e - s - Len(TagName) - 2)
            End If
      End If
End Function

' -----------------------------------------------------------------------------
' GetItemFromPlottersCacheCol (Function), source lines 88-94
' -----------------------------------------------------------------------------
Private Function GetItemFromPlottersCacheCol(Key As String) As Plotter2DModelCache
    On Error GoTo 10
    Set GetItemFromPlottersCacheCol = PlottersCacheCol(Key)
    Exit Function
10
    Set GetItemFromPlottersCacheCol = Nothing
End Function

' -----------------------------------------------------------------------------
' GetModel_600 (Function), source lines 96-115
' -----------------------------------------------------------------------------
Public Function GetModel_600() As Plotter2DModel
    If Not StandardPlotter600_2DModel_Initialized Then
        InitStandardPlotter600_2DModel
        
        With StandardPlotter600_2DModel
            Dim boundingBox As RECTL
            boundingBox = CalculatePlotterBoundingBox(.Curves)
            Set .CarriageCurves = GetStandardCarriageCurves
            .PlotterWorkWidth = 600000
            .top = boundingBox.top
            .bottom = boundingBox.bottom
            .left = boundingBox.left
            .right = boundingBox.right
            .Height = .top - .bottom
            .Width = .right - .left
        End With
        
    End If
    GetModel_600 = StandardPlotter600_2DModel
End Function

' -----------------------------------------------------------------------------
' GetModel_1600 (Function), source lines 116-135
' -----------------------------------------------------------------------------
Public Function GetModel_1600() As Plotter2DModel
    If Not StandardPlotter1600_2DModel_Initialized Then
        InitStandardPlotter1600_2DModel
        
        With StandardPlotter1600_2DModel
            Dim boundingBox As RECTL
            boundingBox = CalculatePlotterBoundingBox(.Curves)
            Set .CarriageCurves = GetStandardCarriageCurves
            .PlotterWorkWidth = 1600000
            .top = boundingBox.top
            .bottom = boundingBox.bottom
            .left = boundingBox.left
            .right = boundingBox.right
            .Height = .top - .bottom
            .Width = .right - .left
        End With
        
    End If
    GetModel_1600 = StandardPlotter1600_2DModel
End Function

' -----------------------------------------------------------------------------
' GetStandardCarriageCurves (Function), source lines 136-139
' -----------------------------------------------------------------------------
Private Function GetStandardCarriageCurves() As Collection
    If Not StandardCarriageCurvesInitialized Then InitStandardCarriage
    Set GetStandardCarriageCurves = StandardCarriageCurves
End Function

' -----------------------------------------------------------------------------
' GetAxisCurves (Function), source lines 140-143
' -----------------------------------------------------------------------------
Public Function GetAxisCurves() As Collection
    If Not AxisCurves_Initialized Then InitAxisCurves
    Set GetAxisCurves = AxisCurves
End Function

' -----------------------------------------------------------------------------
' GetLogoCurves (Function), source lines 144-147
' -----------------------------------------------------------------------------
Public Function GetLogoCurves() As Collection
    If Not LogoCurvesInitialized Then InitLogoCurves
    Set GetLogoCurves = LogoCurves
End Function

' -----------------------------------------------------------------------------
' CalculatePlotterBoundingBox (Function), source lines 148-173
' -----------------------------------------------------------------------------
Private Function CalculatePlotterBoundingBox(Curves As Collection) As RECTL
    Dim q As Long
    Dim p As waPath
    
    Dim top As Long: top = 2147483647
    Dim left As Long: left = 2147483647
    Dim right As Long: right = -2147483647
    Dim bottom As Long: bottom = -2147483647
    
    For q = 1 To Curves.count
        Set p = Curves(q)
        With p.Bounds
            If .top < top Then top = .top
            If .bottom > bottom Then bottom = .bottom
            If .right > right Then right = .right
            If .left < left Then left = .left
        End With
    Next
    With CalculatePlotterBoundingBox
        .top = bottom
        .bottom = top
        .left = left
        .right = right
    End With
    
End Function

' -----------------------------------------------------------------------------
' InitStandardCarriage (Sub), source lines 175-563
' -----------------------------------------------------------------------------
Private Sub InitStandardCarriage()
    Dim MyCurves As New Collection
    Set StandardCarriageCurves = MyCurves
    Dim path As waPath

    Set path = New waPath
    With path
        .StartSubPath
        .AddLineSegment -290, 78872
        .AddLineSegment -290, 25425
        .AddLineSegment -65642, 25425
        .AddLineSegment -65642, 78872
        .AddLineSegment -290, 78872
        .CloseSubPath

        Set .brush = New waBrush
        .brush.SetColor RGB(91, 91, 91), 100
    End With
    MyCurves.Add path

    Set path = New waPath
    With path
        .StartSubPath
        .AddLineSegment -14577, 11602
        .AddLineSegment -14577, 54464
        .AddBezierSegment -14577, 54464, -12146, 53249, -8623, 49286, -7697, 45601
        .AddLineSegment -1744, 11602
        .AddLineSegment -14577, 11602
        .CloseSubPath

        Set .brush = New waBrush
        .brush.SetColor RGB(67, 66, 66), 100
    End With
    MyCurves.Add path

    Set path = New waPath
    With path
        .StartSubPath
        .AddLineSegment -60879, 17555
        .AddLineSegment -17090, 17555
        .AddBezierSegment -17090, 17555, -15708, 17555, -14577, 18686, -14577, 20069
        .AddLineSegment -14577, 61211
        .AddBezierSegment -14577, 61211, -14577, 62594, -15708, 63725, -17090, 63725
        .AddLineSegment -60879, 63725
        .AddBezierSegment -60879, 63725, -62261, 63725, -63392, 62594, -63392, 61211
        .AddLineSegment -63392, 20069
        .AddBezierSegment -63392, 20069, -63392, 18686, -62261, 17555, -60879, 17555
        .CloseSubPath

        Set .brush = New waBrush
        .brush.SetColor RGB(43, 42, 41), 100
    End With
    MyCurves.Add path

    Set path = New waPath
    With path
        .StartSubPath
        .AddLineSegment 5831, 11287
        .AddLineSegment 3085, 2654
        .AddBezierSegment 3085, 2654, 2688, 1407, 2013, 273, 704, 273
        .AddLineSegment -621, 273
        .AddBezierSegment -621, 273, -1930, 273, -2605, 1406, -3002, 2654
        .AddLineSegment -5748, 11287
        .AddLineSegment -5748, 35386
        .AddBezierSegment -5748, 35386, -5748, 35956, -5282, 36422, -4712, 36422
        .AddLineSegment 4795, 36422
        .AddBezierSegment 4795, 36422, 5365, 36422, 5831, 35956, 5831, 35386
        .AddLineSegment 5831, 11287
        .CloseSubPath

        Set .brush = New waBrush
        .brush.SetLinearGradient -11000, 6000, 0, 6000, RGB(43, 42, 41), RGB(84, 84, 84), WrapModeTileFlipXY
    End With
    MyCurves.Add path

    Set path = New waPath
    With path
        .StartSubPath
        .AddLineSegment -7696, 11287
        .AddLineSegment 7778, 11287
        .AddBezierSegment 7778, 11287, 8581, 11287, 9237, 11943, 9237, 12745
        .AddLineSegment 9237, 26132
        .AddBezierSegment 9237, 26132, 9237, 26935, 8581, 27591, 7778, 27591
        .AddLineSegment -7696, 27591
        .AddBezierSegment -7696, 27591, -8498, 27591, -9154, 26935, -9154, 26132
        .AddLineSegment -9154, 12745
        .AddBezierSegment -9154, 12745, -9154, 11943, -8498, 11287, -7696, 11287
        .CloseSubPath

        Set .brush = New waBrush
        .brush.SetColor RGB(50, 50, 50), 100
    End With
    MyCurves.Add path

    Set path = New waPath
    With path
        .StartSubPath
        .AddLineSegment -3728, 39035
        .AddLineSegment 3648, 39035
        .AddLineSegment 3648, 36422
        .AddLineSegment -3728, 36422
        .AddLineSegment -3728, 39035
        .CloseSubPath

        Set .brush = New waBrush
        .brush.SetColor RGB(43, 42, 41), 100
    End With
    MyCurves.Add path

    Set path = New waPath
    With path
        .StartSubPath
        .AddLineSegment -5348, 39035
        .AddLineSegment 5269, 39035
        .AddBezierSegment 5269, 39035, 5505, 39035, 5698, 39228, 5698, 39465
        .AddLineSegment 5698, 43069
        .AddBezierSegment 5698, 43069, 5698, 43306, 5505, 43499, 5269, 43499
        .AddLineSegment -5348, 43499
        .AddBezierSegment -5348, 43499, -5585, 43499, -5778, 43306, -5778, 43069
        .AddLineSegment -5778, 39465
        .AddBezierSegment -5778, 39465, -5778, 39228, -5585, 39035, -5348, 39035
        .CloseSubPath

        Set .brush = New waBrush
        .brush.SetColor RGB(67, 66, 66), 100
    End With
    MyCurves.Add path

    Set path = New waPath
    With path
        .StartSubPath
        .AddLineSegment -801, 49320
        .AddLineSegment 722, 49320
        .AddBezierSegment 722, 49320, 821, 49320, 903, 49238, 903, 49138
        .AddLineSegment 903, 43681
        .AddBezierSegment 903, 43681, 903, 43581, 821, 43499, 722, 43499
        .AddLineSegment -801, 43499
        .AddBezierSegment -801, 43499, -901, 43499, -983, 43581, -983, 43681
        .AddLineSegment -983, 49138
        .AddBezierSegment -983, 49138, -983, 49238, -901, 49320, -801, 49320
        .CloseSubPath

        Set .brush = New waBrush
        .brush.SetLinearGradient -2000, 46000, 0, 46000, RGB(178, 179, 179), RGB(254, 254, 254), WrapModeTileFlipXY
    End With
    MyCurves.Add path

    Set path = New waPath
    With path
        .StartSubPath
        .AddLineSegment -7696, 11287
        .AddLineSegment -5782, 11287
        .AddBezierSegment -5782, 11287, -4979, 11287, -4323, 11943, -4323, 12745
        .AddLineSegment -4323, 26132
        .AddBezierSegment -4323, 26132, -4323, 26935, -4979, 27591, -5782, 27591
        .AddLineSegment -7696, 27591
        .AddBezierSegment -7696, 27591, -8498, 27591, -9154, 26934, -9154, 26132
        .AddLineSegment -9154, 12745
        .AddBezierSegment -9154, 12745, -9154, 11943, -8498, 11287, -7696, 11287
        .CloseSubPath

        Set .brush = New waBrush
        .brush.SetLinearGradient -12000, 19000, -7000, 19000, RGB(31, 27, 32), RGB(77, 79, 77), WrapModeTileFlipXY
    End With
    MyCurves.Add path

    Set path = New waPath
    With path
        .StartSubPath
        .AddLineSegment 56, 37407
        .AddBezierSegment 56, 37407, 380, 37407, 642, 37669, 642, 37993
        .AddBezierSegment 642, 37993, 642, 38317, 380, 38579, 56, 38579
        .AddBezierSegment 56, 38579, -268, 38579, -530, 38317, -530, 37993
        .AddBezierSegment -530, 37993, -530, 37669, -268, 37407, 56, 37407
        .CloseSubPath

        Set .brush = New waBrush
        .brush.SetColor RGB(178, 179, 179), 100
    End With
    MyCurves.Add path

    Set path = New waPath
    With path
        .StartSubPath
        .AddLineSegment -2373, 36422
        .AddLineSegment -1959, 36422
        .AddLineSegment -1959, 33005
        .AddLineSegment -2373, 33005
        .AddLineSegment -2373, 36422
        .CloseSubPath

        Set .brush = New waBrush
        .brush.SetColor RGB(178, 179, 179), 100
    End With
    MyCurves.Add path

    Set path = New waPath
    With path
        .StartSubPath
        .AddLineSegment 1474, 36422
        .AddLineSegment 1887, 36422
        .AddLineSegment 1887, 33005
        .AddLineSegment 1474, 33005
        .AddLineSegment 1474, 36422
        .CloseSubPath

        Set .brush = New waBrush
        .brush.SetColor RGB(178, 179, 179), 100
    End With
    MyCurves.Add path

    Set path = New waPath
    With path
        .StartSubPath
        .AddLineSegment -247, 42975
        .AddLineSegment 167, 42975
        .AddLineSegment 167, 39558
        .AddLineSegment -247, 39558
        .AddLineSegment -247, 42975
        .CloseSubPath

        Set .brush = New waBrush
        .brush.SetColor RGB(43, 42, 41), 100
    End With
    MyCurves.Add path

    Set path = New waPath
    With path
        .StartSubPath
        .AddLineSegment -1465, 42975
        .AddLineSegment -1052, 42975
        .AddLineSegment -1052, 39558
        .AddLineSegment -1465, 39558
        .AddLineSegment -1465, 42975
        .CloseSubPath

        Set .brush = New waBrush
        .brush.SetColor RGB(43, 42, 41), 100
    End With
    MyCurves.Add path

    Set path = New waPath
    With path
        .StartSubPath
        .AddLineSegment 1057, 42975
        .AddLineSegment 1470, 42975
        .AddLineSegment 1470, 39558
        .AddLineSegment 1057, 39558
        .AddLineSegment 1057, 42975
        .CloseSubPath

        Set .brush = New waBrush
        .brush.SetColor RGB(43, 42, 41), 100
    End With
    MyCurves.Add path

    Set path = New waPath
    With path
        .StartSubPath
        .AddLineSegment 3339, 42975
        .AddLineSegment 3752, 42975
        .AddLineSegment 3752, 39558
        .AddLineSegment 3339, 39558
        .AddLineSegment 3339, 42975
        .CloseSubPath

        Set .brush = New waBrush
        .brush.SetColor RGB(43, 42, 41), 100
    End With
    MyCurves.Add path

    Set path = New waPath
    With path
        .StartSubPath
        .AddLineSegment 2120, 42975
        .AddLineSegment 2534, 42975
        .AddLineSegment 2534, 39558
        .AddLineSegment 2120, 39558
        .AddLineSegment 2120, 42975
        .CloseSubPath

        Set .brush = New waBrush
        .brush.SetColor RGB(43, 42, 41), 100
    End With
    MyCurves.Add path

    Set path = New waPath
    With path
        .StartSubPath
        .AddLineSegment 4642, 42975
        .AddLineSegment 5056, 42975
        .AddLineSegment 5056, 39558
        .AddLineSegment 4642, 39558
        .AddLineSegment 4642, 42975
        .CloseSubPath

        Set .brush = New waBrush
        .brush.SetColor RGB(43, 42, 41), 100
    End With
    MyCurves.Add path

    Set path = New waPath
    With path
        .StartSubPath
        .AddLineSegment -3904, 42975
        .AddLineSegment -3491, 42975
        .AddLineSegment -3491, 39558
        .AddLineSegment -3904, 39558
        .AddLineSegment -3904, 42975
        .CloseSubPath

        Set .brush = New waBrush
        .brush.SetColor RGB(43, 42, 41), 100
    End With
    MyCurves.Add path

    Set path = New waPath
    With path
        .StartSubPath
        .AddLineSegment -5123, 42975
        .AddLineSegment -4709, 42975
        .AddLineSegment -4709, 39558
        .AddLineSegment -5123, 39558
        .AddLineSegment -5123, 42975
        .CloseSubPath

        Set .brush = New waBrush
        .brush.SetColor RGB(43, 42, 41), 100
    End With
    MyCurves.Add path

    Set path = New waPath
    With path
        .StartSubPath
        .AddLineSegment -2601, 42975
        .AddLineSegment -2187, 42975
        .AddLineSegment -2187, 39558
        .AddLineSegment -2601, 39558
        .AddLineSegment -2601, 42975
        .CloseSubPath

        Set .brush = New waBrush
        .brush.SetColor RGB(43, 42, 41), 100
    End With
    MyCurves.Add path

    Set path = New waPath
    With path
        .StartSubPath
        .AddLineSegment -47098, 37127
        .AddLineSegment -43167, 37517
        .AddBezierSegment -43167, 37517, -42930, 36202, -42456, 35240, -41736, 34622
        .AddBezierSegment -41736, 34622, -41015, 34004, -40044, 33697, -38822, 33697
        .AddBezierSegment -38822, 33697, -37530, 33697, -36554, 33971, -35899, 34520
        .AddBezierSegment -35899, 34520, -35244, 35068, -34914, 35709, -34914, 36444
        .AddBezierSegment -34914, 36444, -34914, 36918, -35053, 37317, -35327, 37652
        .AddBezierSegment -35327, 37652, -35601, 37982, -36085, 38270, -36773, 38512
        .AddBezierSegment -36773, 38512, -37242, 38679, -38311, 38967, -39984, 39385
        .AddBezierSegment -39984, 39385, -42135, 39924, -43645, 40584, -44514, 41370
        .AddBezierSegment -44514, 41370, -45737, 42471, -46346, 43814, -46346, 45399
        .AddBezierSegment -46346, 45399, -46346, 46416, -46057, 47374, -45486, 48261
        .AddBezierSegment -45486, 48261, -44910, 49149, -44082, 49823, -43000, 50287
        .AddBezierSegment -43000, 50287, -41917, 50752, -40616, 50984, -39082, 50984
        .AddBezierSegment -39082, 50984, -36587, 50984, -34705, 50431, -33445, 49321
        .AddBezierSegment -33445, 49321, -32181, 48210, -31521, 46732, -31456, 44878
        .AddLineSegment -35499, 44739
        .AddBezierSegment -35499, 44739, -35671, 45770, -36038, 46509, -36601, 46960
        .AddBezierSegment -36601, 46960, -37163, 47411, -38009, 47639, -39133, 47639
        .AddBezierSegment -39133, 47639, -40295, 47639, -41206, 47397, -41861, 46914
        .AddBezierSegment -41861, 46914, -42284, 46602, -42498, 46189, -42498, 45668
        .AddBezierSegment -42498, 45668, -42498, 45194, -42298, 44790, -41898, 44455
        .AddBezierSegment -41898, 44455, -41392, 44023, -40160, 43577, -38204, 43112
        .AddBezierSegment -38204, 43112, -36247, 42652, -34802, 42169, -33868, 41676
        .AddBezierSegment -33868, 41676, -32929, 41179, -32195, 40505, -31665, 39646
        .AddBezierSegment -31665, 39646, -31136, 38786, -30871, 37726, -30871, 36467
        .AddBezierSegment -30871, 36467, -30871, 35324, -31187, 34250, -31823, 33251
        .AddBezierSegment -31823, 33251, -32460, 32252, -33357, 31509, -34519, 31025
        .AddBezierSegment -34519, 31025, -35680, 30537, -37130, 30296, -38864, 30296
        .AddBezierSegment -38864, 30296, -41387, 30296, -43325, 30877, -44677, 32043
        .AddBezierSegment -44677, 32043, -46029, 33205, -46833, 34901, -47098, 37127
        .CloseSubPath

        Set .brush = New waBrush
        .brush.SetColor RGB(235, 236, 236), 100
    End With
    MyCurves.Add path

    StandardCarriageCurvesInitialized = True
End Sub

' -----------------------------------------------------------------------------
' InitStandardPlotter600_2DModel (Sub), source lines 565-2731
' -----------------------------------------------------------------------------
Private Sub InitStandardPlotter600_2DModel()

    Dim MyCurves As New Collection
    Set StandardPlotter600_2DModel.Curves = MyCurves
    Dim path As waPath

    Set path = New waPath
    With path
        .StartSubPath
        .AddLineSegment 70178, -770784
        .AddLineSegment 69470, -756279
        .AddLineSegment 88573, -756279
        .AddLineSegment 87938, -770784
        .AddBezierSegment 87938, -770784, 84312, -771828, 73671, -772120, 70178, -770784
        .CloseSubPath

        Set .brush = New waBrush
        .brush.SetLinearGradient 60000, -764000, 79000, -764000, RGB(67, 66, 66), RGB(143, 143, 143), WrapModeTileFlipXY
    End With
    MyCurves.Add path

    Set path = New waPath
    With path
        .StartSubPath
        .AddLineSegment -14912, -856615
        .AddLineSegment -15620, -842109
        .AddLineSegment 3483, -842109
        .AddLineSegment 2848, -856615
        .AddBezierSegment 2848, -856615, -778, -857659, -11419, -857951, -14912, -856615
        .CloseSubPath

        Set .brush = New waBrush
        .brush.SetLinearGradient -25000, -850000, -6000, -850000, RGB(67, 66, 66), RGB(143, 143, 143), WrapModeTileFlipXY
    End With
    MyCurves.Add path

    Set path = New waPath
    With path
        .StartSubPath
        .AddLineSegment 66493, -749902
        .AddLineSegment -22804, -839200
        .AddBezierSegment -22804, -839200, -24185, -840581, -25043, -842486, -25043, -844582
        .AddLineSegment -25043, -848815
        .AddBezierSegment -25043, -848815, -25043, -853006, -21614, -856435, -17423, -856435
        .AddLineSegment -11708, -856435
        .AddLineSegment -5993, -856435
        .AddBezierSegment -5993, -856435, -3897, -856435, -1992, -855577, -611, -854197
        .AddLineSegment 88686, -764899
        .AddBezierSegment 88686, -764899, 90067, -763518, 90924, -761613, 90924, -759517
        .AddLineSegment 90924, -755284
        .AddBezierSegment 90924, -755284, 90924, -751093, 87495, -747664, 83304, -747664
        .AddLineSegment 77589, -747664
        .AddLineSegment 71874, -747664
        .AddBezierSegment 71874, -747664, 69779, -747664, 67874, -748521, 66493, -749902
        .CloseSubPath

        Set .brush = New waBrush
        .brush.SetLinearGradient -9000, -691000, 75000, -765000, RGB(164, 185, 195), RGB(218, 218, 218), WrapModeTileFlipXY
    End With
    MyCurves.Add path

    Set path = New waPath
    With path
        .StartSubPath
        .AddLineSegment -17740, -854711
        .AddLineSegment -6310, -854711
        .AddBezierSegment -6310, -854711, -3223, -854711, -690, -852178, -690, -849091
        .AddLineSegment -690, -844858
        .AddBezierSegment -690, -844858, -690, -841771, -3223, -839238, -6310, -839238
        .AddLineSegment -17740, -839238
        .AddBezierSegment -17740, -839238, -20827, -839238, -23360, -841771, -23360, -844858
        .AddLineSegment -23360, -849091
        .AddBezierSegment -23360, -849091, -23360, -852178, -20827, -854711, -17740, -854711
        .CloseSubPath

        Set .brush = New waBrush
        .brush.SetLinearGradient -20000, -835000, -4000, -846000, RGB(114, 114, 113), RGB(189, 182, 182), WrapModeTileFlipXY
    End With
    MyCurves.Add path

    Set path = New waPath
    With path
        .StartSubPath
        .AddLineSegment -584783, -770784
        .AddLineSegment -585491, -756279
        .AddLineSegment -566388, -756279
        .AddLineSegment -567023, -770784
        .AddBezierSegment -567023, -770784, -570648, -771828, -581290, -772120, -584783, -770784
        .CloseSubPath

        Set .brush = New waBrush
        .brush.SetLinearGradient -595000, -764000, -576000, -764000, RGB(67, 66, 66), RGB(143, 143, 143), WrapModeTileFlipXY
    End With
    MyCurves.Add path

    Set path = New waPath
    With path
        .StartSubPath
        .AddLineSegment -669873, -856615
        .AddLineSegment -670581, -842109
        .AddLineSegment -651478, -842109
        .AddLineSegment -652113, -856615
        .AddBezierSegment -652113, -856615, -655739, -857659, -666380, -857951, -669873, -856615
        .CloseSubPath

        Set .brush = New waBrush
        .brush.SetLinearGradient -680000, -850000, -661000, -850000, RGB(67, 66, 66), RGB(143, 143, 143), WrapModeTileFlipXY
    End With
    MyCurves.Add path

    Set path = New waPath
    With path
        .StartSubPath
        .AddLineSegment -588468, -749902
        .AddLineSegment -677765, -839200
        .AddBezierSegment -677765, -839200, -679146, -840581, -680004, -842486, -680004, -844582
        .AddLineSegment -680004, -848815
        .AddBezierSegment -680004, -848815, -680004, -853006, -676574, -856435, -672384, -856435
        .AddLineSegment -666668, -856435
        .AddLineSegment -660953, -856435
        .AddBezierSegment -660953, -856435, -658858, -856435, -656953, -855577, -655572, -854197
        .AddLineSegment -566275, -764899
        .AddBezierSegment -566275, -764899, -564894, -763518, -564037, -761613, -564037, -759517
        .AddLineSegment -564037, -755284
        .AddBezierSegment -564037, -755284, -564037, -751093, -567466, -747664, -571657, -747664
        .AddLineSegment -577372, -747664
        .AddLineSegment -583087, -747664
        .AddBezierSegment -583087, -747664, -585182, -747664, -587087, -748521, -588468, -749902
        .CloseSubPath

        Set .brush = New waBrush
        .brush.SetLinearGradient -664000, -691000, -580000, -765000, RGB(164, 185, 195), RGB(218, 218, 218), WrapModeTileFlipXY
    End With
    MyCurves.Add path

    Set path = New waPath
    With path
        .StartSubPath
        .AddLineSegment 14175, -401920
        .AddLineSegment 39662, -401920
        .AddLineSegment 39662, -380859
        .AddLineSegment 14175, -380859
        .AddLineSegment 14175, -401920
        .CloseSubPath

        Set .brush = New waBrush
        .brush.SetLinearGradient 27000, -412000, 27000, -434000, RGB(43, 42, 41), RGB(141, 141, 141), WrapModeTileFlipXY
    End With
    MyCurves.Add path

    Set path = New waPath
    With path
        .StartSubPath
        .AddLineSegment -639574, -401920
        .AddLineSegment -614088, -401920
        .AddLineSegment -614088, -380859
        .AddLineSegment -639574, -380859
        .AddLineSegment -639574, -401920
        .CloseSubPath

        Set .brush = New waBrush
        .brush.SetLinearGradient -627000, -412000, -627000, -434000, RGB(43, 42, 41), RGB(141, 141, 141), WrapModeTileFlipXY
    End With
    MyCurves.Add path

    Set path = New waPath
    With path
        .StartSubPath
        .AddLineSegment -636630, -57346
        .AddLineSegment -618003, -57346
        .AddLineSegment -618003, -798107
        .AddLineSegment -636630, -798107
        .AddLineSegment -636630, -57346
        .CloseSubPath

        Set .brush = New waBrush
        .brush.SetLinearGradient -646000, -428000, -627000, -428000, RGB(164, 185, 195), RGB(208, 206, 210), WrapModeTileFlipXY
    End With
    MyCurves.Add path

    Set path = New waPath
    With path
        .StartSubPath
        .AddLineSegment 16091, -434570
        .AddLineSegment 16091, -463287
        .AddLineSegment -616003, -463287
        .AddLineSegment -616003, -434570
        .AddLineSegment 16091, -434570
        .CloseSubPath

        Set .brush = New waBrush
        .brush.SetLinearGradient -300000, -412000, -300000, -441000, RGB(164, 185, 195), RGB(214, 211, 218), WrapModeTileFlipXY
    End With
    MyCurves.Add path

    Set path = New waPath
    With path
        .StartSubPath
        .AddLineSegment 18091, -57346
        .AddLineSegment 36718, -57346
        .AddLineSegment 36718, -798107
        .AddLineSegment 18091, -798107
        .AddLineSegment 18091, -57346
        .CloseSubPath

        Set .brush = New waBrush
        .brush.SetLinearGradient 9000, -428000, 27000, -428000, RGB(164, 185, 195), RGB(208, 206, 210), WrapModeTileFlipXY
    End With
    MyCurves.Add path

    Set path = New waPath
    With path
        .StartSubPath
        .AddLineSegment 16091, -494084
        .AddLineSegment 16091, -522801
        .AddLineSegment -616003, -522801
        .AddLineSegment -616003, -494084
        .AddLineSegment 16091, -494084
        .CloseSubPath

        Set .brush = New waBrush
        .brush.SetLinearGradient -300000, -472000, -300000, -501000, RGB(164, 185, 195), RGB(214, 211, 218), WrapModeTileFlipXY
    End With
    MyCurves.Add path

    Set path = New waPath
    With path
        .StartSubPath
        .AddLineSegment 14175, -382007
        .AddLineSegment 14175, -400634
        .AddLineSegment -614088, -400634
        .AddLineSegment -614088, -382007
        .AddLineSegment 14175, -382007
        .CloseSubPath

        Set .brush = New waBrush
        .brush.SetLinearGradient -300000, -373000, -300000, -391000, RGB(164, 185, 195), RGB(208, 206, 210), WrapModeTileFlipXY
    End With
    MyCurves.Add path

    Set path = New waPath
    With path
        .StartSubPath
        .AddLineSegment -618003, -539769
        .AddLineSegment -616003, -539769
        .AddLineSegment -616003, -409805
        .AddLineSegment -618003, -409805
        .AddLineSegment -618003, -539769
        .CloseSubPath

        Set .brush = New waBrush
        .brush.SetColor RGB(142, 164, 174), 100
    End With
    MyCurves.Add path

    Set path = New waPath
    With path
        .StartSubPath
        .AddLineSegment 16091, -539769
        .AddLineSegment 18091, -539769
        .AddLineSegment 18091, -409805
        .AddLineSegment 16091, -409805
        .AddLineSegment 16091, -539769
        .CloseSubPath

        Set .brush = New waBrush
        .brush.SetColor RGB(142, 164, 174), 100
    End With
    MyCurves.Add path

    Set path = New waPath
    With path
        .StartSubPath
        .AddLineSegment 36718, -508548
        .AddLineSegment 38718, -508548
        .AddLineSegment 38718, -409805
        .AddLineSegment 36718, -409805
        .AddLineSegment 36718, -508548
        .CloseSubPath

        Set .brush = New waBrush
        .brush.SetColor RGB(142, 164, 174), 100
    End With
    MyCurves.Add path

    Set path = New waPath
    With path
        .StartSubPath
        .AddLineSegment -638630, -508548
        .AddLineSegment -636630, -508548
        .AddLineSegment -636630, -409805
        .AddLineSegment -638630, -409805
        .AddLineSegment -638630, -508548
        .CloseSubPath

        Set .brush = New waBrush
        .brush.SetColor RGB(142, 164, 174), 100
    End With
    MyCurves.Add path

    Set path = New waPath
    With path
        .StartSubPath
        .AddLineSegment -672701, -854711
        .AddLineSegment -661271, -854711
        .AddBezierSegment -661271, -854711, -658184, -854711, -655651, -852178, -655651, -849091
        .AddLineSegment -655651, -844858
        .AddBezierSegment -655651, -844858, -655651, -841771, -658184, -839238, -661271, -839238
        .AddLineSegment -672701, -839238
        .AddBezierSegment -672701, -839238, -675788, -839238, -678321, -841771, -678321, -844858
        .AddLineSegment -678321, -849091
        .AddBezierSegment -678321, -849091, -678321, -852178, -675788, -854711, -672701, -854711
        .CloseSubPath

        Set .brush = New waBrush
        .brush.SetLinearGradient -675000, -835000, -659000, -846000, RGB(114, 114, 113), RGB(189, 182, 182), WrapModeTileFlipXY
    End With
    MyCurves.Add path

    Set path = New waPath
    With path
        .StartSubPath
        .AddLineSegment -670456, -57346
        .AddLineSegment 70541, -57346
        .AddLineSegment 70541, -25346
        .AddLineSegment -670456, -25346
        .AddLineSegment -670456, -57346
        .CloseSubPath

        Set .brush = New waBrush
        .brush.SetColor RGB(164, 185, 195), 100
    End With
    MyCurves.Add path

    Set path = New waPath
    With path
        .StartSubPath
        .AddLineSegment -670456, -17824
        .AddLineSegment 70541, -17824
        .AddLineSegment 70541, -25346
        .AddLineSegment -670456, -25346
        .AddLineSegment -670456, -17824
        .CloseSubPath

        Set .brush = New waBrush
        .brush.SetColor RGB(254, 254, 254), 100
    End With
    MyCurves.Add path

    Set path = New waPath
    With path
        .StartSubPath
        .AddLineSegment -645456, -25346
        .AddLineSegment -644456, -25346
        .AddLineSegment -644456, -21737
        .AddLineSegment -645456, -21737
        .AddLineSegment -645456, -25346
        .CloseSubPath

        .StartSubPath
        .AddLineSegment -5459, -25346
        .AddLineSegment -4459, -25346
        .AddLineSegment -4459, -21737
        .AddLineSegment -5459, -21737
        .AddLineSegment -5459, -25346
        .CloseSubPath

        .StartSubPath
        .AddLineSegment -459, -25346
        .AddLineSegment 541, -25346
        .AddLineSegment 541, -17824
        .AddLineSegment -459, -17824
        .AddLineSegment -459, -25346
        .CloseSubPath

        .StartSubPath
        .AddLineSegment 39541, -25346
        .AddLineSegment 40541, -25346
        .AddLineSegment 40541, -17824
        .AddLineSegment 39541, -17824
        .AddLineSegment 39541, -25346
        .CloseSubPath

        .StartSubPath
        .AddLineSegment 49541, -25346
        .AddLineSegment 50541, -25346
        .AddLineSegment 50541, -17824
        .AddLineSegment 49541, -17824
        .AddLineSegment 49541, -25346
        .CloseSubPath

        .StartSubPath
        .AddLineSegment 29541, -25346
        .AddLineSegment 30541, -25346
        .AddLineSegment 30541, -17824
        .AddLineSegment 29541, -17824
        .AddLineSegment 29541, -25346
        .CloseSubPath

        .StartSubPath
        .AddLineSegment 19541, -25346
        .AddLineSegment 20541, -25346
        .AddLineSegment 20541, -17824
        .AddLineSegment 19541, -17824
        .AddLineSegment 19541, -25346
        .CloseSubPath

        .StartSubPath
        .AddLineSegment 9541, -25346
        .AddLineSegment 10541, -25346
        .AddLineSegment 10541, -17824
        .AddLineSegment 9541, -17824
        .AddLineSegment 9541, -25346
        .CloseSubPath

        .StartSubPath
        .AddLineSegment -10459, -25346
        .AddLineSegment -9459, -25346
        .AddLineSegment -9459, -17824
        .AddLineSegment -10459, -17824
        .AddLineSegment -10459, -25346
        .CloseSubPath

        .StartSubPath
        .AddLineSegment 34541, -25346
        .AddLineSegment 35541, -25346
        .AddLineSegment 35541, -21737
        .AddLineSegment 34541, -21737
        .AddLineSegment 34541, -25346
        .CloseSubPath

        .StartSubPath
        .AddLineSegment 54541, -25346
        .AddLineSegment 55541, -25346
        .AddLineSegment 55541, -21737
        .AddLineSegment 54541, -21737
        .AddLineSegment 54541, -25346
        .CloseSubPath

        .StartSubPath
        .AddLineSegment 44541, -25346
        .AddLineSegment 45541, -25346
        .AddLineSegment 45541, -21737
        .AddLineSegment 44541, -21737
        .AddLineSegment 44541, -25346
        .CloseSubPath

        .StartSubPath
        .AddLineSegment 24541, -25346
        .AddLineSegment 25541, -25346
        .AddLineSegment 25541, -21737
        .AddLineSegment 24541, -21737
        .AddLineSegment 24541, -25346
        .CloseSubPath

        .StartSubPath
        .AddLineSegment 14541, -25346
        .AddLineSegment 15541, -25346
        .AddLineSegment 15541, -21737
        .AddLineSegment 14541, -21737
        .AddLineSegment 14541, -25346
        .CloseSubPath

        .StartSubPath
        .AddLineSegment 4541, -25346
        .AddLineSegment 5541, -25346
        .AddLineSegment 5541, -21737
        .AddLineSegment 4541, -21737
        .AddLineSegment 4541, -25346
        .CloseSubPath

        .StartSubPath
        .AddLineSegment -15459, -25346
        .AddLineSegment -14459, -25346
        .AddLineSegment -14459, -21737
        .AddLineSegment -15459, -21737
        .AddLineSegment -15459, -25346
        .CloseSubPath

        .StartSubPath
        .AddLineSegment -325457, -25346
        .AddLineSegment -324457, -25346
        .AddLineSegment -324457, -21737
        .AddLineSegment -325457, -21737
        .AddLineSegment -325457, -25346
        .CloseSubPath

        .StartSubPath
        .AddLineSegment -165458, -25346
        .AddLineSegment -164458, -25346
        .AddLineSegment -164458, -21737
        .AddLineSegment -165458, -21737
        .AddLineSegment -165458, -25346
        .CloseSubPath

        .StartSubPath
        .AddLineSegment -160458, -25346
        .AddLineSegment -159458, -25346
        .AddLineSegment -159458, -17824
        .AddLineSegment -160458, -17824
        .AddLineSegment -160458, -25346
        .CloseSubPath

        .StartSubPath
        .AddLineSegment -80458, -25346
        .AddLineSegment -79458, -25346
        .AddLineSegment -79458, -17824
        .AddLineSegment -80458, -17824
        .AddLineSegment -80458, -25346
        .CloseSubPath

        .StartSubPath
        .AddLineSegment -40458, -25346
        .AddLineSegment -39458, -25346
        .AddLineSegment -39458, -17824
        .AddLineSegment -40458, -17824
        .AddLineSegment -40458, -25346
        .CloseSubPath

        .StartSubPath
        .AddLineSegment -20459, -25346
        .AddLineSegment -19459, -25346
        .AddLineSegment -19459, -17824
        .AddLineSegment -20459, -17824
        .AddLineSegment -20459, -25346
        .CloseSubPath

        .StartSubPath
        .AddLineSegment -30458, -25346
        .AddLineSegment -29458, -25346
        .AddLineSegment -29458, -17824
        .AddLineSegment -30458, -17824
        .AddLineSegment -30458, -25346
        .CloseSubPath

        .StartSubPath
        .AddLineSegment -50458, -25346
        .AddLineSegment -49458, -25346
        .AddLineSegment -49458, -17824
        .AddLineSegment -50458, -17824
        .AddLineSegment -50458, -25346
        .CloseSubPath

        .StartSubPath
        .AddLineSegment -60458, -25346
        .AddLineSegment -59458, -25346
        .AddLineSegment -59458, -17824
        .AddLineSegment -60458, -17824
        .AddLineSegment -60458, -25346
        .CloseSubPath

        .StartSubPath
        .AddLineSegment -70458, -25346
        .AddLineSegment -69458, -25346
        .AddLineSegment -69458, -17824
        .AddLineSegment -70458, -17824
        .AddLineSegment -70458, -25346
        .CloseSubPath

        .StartSubPath
        .AddLineSegment -90458, -25346
        .AddLineSegment -89458, -25346
        .AddLineSegment -89458, -17824
        .AddLineSegment -90458, -17824
        .AddLineSegment -90458, -25346
        .CloseSubPath

        .StartSubPath
        .AddLineSegment -120458, -25346
        .AddLineSegment -119458, -25346
        .AddLineSegment -119458, -17824
        .AddLineSegment -120458, -17824
        .AddLineSegment -120458, -25346
        .CloseSubPath

        .StartSubPath
        .AddLineSegment -100458, -25346
        .AddLineSegment -99458, -25346
        .AddLineSegment -99458, -17824
        .AddLineSegment -100458, -17824
        .AddLineSegment -100458, -25346
        .CloseSubPath

        .StartSubPath
        .AddLineSegment -110458, -25346
        .AddLineSegment -109458, -25346
        .AddLineSegment -109458, -17824
        .AddLineSegment -110458, -17824
        .AddLineSegment -110458, -25346
        .CloseSubPath

        .StartSubPath
        .AddLineSegment -130458, -25346
        .AddLineSegment -129458, -25346
        .AddLineSegment -129458, -17824
        .AddLineSegment -130458, -17824
        .AddLineSegment -130458, -25346
        .CloseSubPath

        .StartSubPath
        .AddLineSegment -140458, -25346
        .AddLineSegment -139458, -25346
        .AddLineSegment -139458, -17824
        .AddLineSegment -140458, -17824
        .AddLineSegment -140458, -25346
        .CloseSubPath

        .StartSubPath
        .AddLineSegment -150458, -25346
        .AddLineSegment -149458, -25346
        .AddLineSegment -149458, -17824
        .AddLineSegment -150458, -17824
        .AddLineSegment -150458, -25346
        .CloseSubPath

        .StartSubPath
        .AddLineSegment -170458, -25346
        .AddLineSegment -169458, -25346
        .AddLineSegment -169458, -17824
        .AddLineSegment -170458, -17824
        .AddLineSegment -170458, -25346
        .CloseSubPath

        .StartSubPath
        .AddLineSegment -85458, -25346
        .AddLineSegment -84458, -25346
        .AddLineSegment -84458, -21737
        .AddLineSegment -85458, -21737
        .AddLineSegment -85458, -25346
        .CloseSubPath

        .StartSubPath
        .AddLineSegment -45458, -25346
        .AddLineSegment -44458, -25346
        .AddLineSegment -44458, -21737
        .AddLineSegment -45458, -21737
        .AddLineSegment -45458, -25346
        .CloseSubPath

        .StartSubPath
        .AddLineSegment -25459, -25346
        .AddLineSegment -24459, -25346
        .AddLineSegment -24459, -21737
        .AddLineSegment -25459, -21737
        .AddLineSegment -25459, -25346
        .CloseSubPath

        .StartSubPath
        .AddLineSegment -35458, -25346
        .AddLineSegment -34458, -25346
        .AddLineSegment -34458, -21737
        .AddLineSegment -35458, -21737
        .AddLineSegment -35458, -25346
        .CloseSubPath

        .StartSubPath
        .AddLineSegment -55458, -25346
        .AddLineSegment -54458, -25346
        .AddLineSegment -54458, -21737
        .AddLineSegment -55458, -21737
        .AddLineSegment -55458, -25346
        .CloseSubPath

        .StartSubPath
        .AddLineSegment -65458, -25346
        .AddLineSegment -64458, -25346
        .AddLineSegment -64458, -21737
        .AddLineSegment -65458, -21737
        .AddLineSegment -65458, -25346
        .CloseSubPath

        .StartSubPath
        .AddLineSegment -75458, -25346
        .AddLineSegment -74458, -25346
        .AddLineSegment -74458, -21737
        .AddLineSegment -75458, -21737
        .AddLineSegment -75458, -25346
        .CloseSubPath

        .StartSubPath
        .AddLineSegment -95458, -25346
        .AddLineSegment -94458, -25346
        .AddLineSegment -94458, -21737
        .AddLineSegment -95458, -21737
        .AddLineSegment -95458, -25346
        .CloseSubPath

        .StartSubPath
        .AddLineSegment -125458, -25346
        .AddLineSegment -124458, -25346
        .AddLineSegment -124458, -21737
        .AddLineSegment -125458, -21737
        .AddLineSegment -125458, -25346
        .CloseSubPath

        .StartSubPath
        .AddLineSegment -105458, -25346
        .AddLineSegment -104458, -25346
        .AddLineSegment -104458, -21737
        .AddLineSegment -105458, -21737
        .AddLineSegment -105458, -25346
        .CloseSubPath

        .StartSubPath
        .AddLineSegment -115458, -25346
        .AddLineSegment -114458, -25346
        .AddLineSegment -114458, -21737
        .AddLineSegment -115458, -21737
        .AddLineSegment -115458, -25346
        .CloseSubPath

        .StartSubPath
        .AddLineSegment -135458, -25346
        .AddLineSegment -134458, -25346
        .AddLineSegment -134458, -21737
        .AddLineSegment -135458, -21737
        .AddLineSegment -135458, -25346
        .CloseSubPath

        .StartSubPath
        .AddLineSegment -145458, -25346
        .AddLineSegment -144458, -25346
        .AddLineSegment -144458, -21737
        .AddLineSegment -145458, -21737
        .AddLineSegment -145458, -25346
        .CloseSubPath

        .StartSubPath
        .AddLineSegment -155458, -25346
        .AddLineSegment -154458, -25346
        .AddLineSegment -154458, -21737
        .AddLineSegment -155458, -21737
        .AddLineSegment -155458, -25346
        .CloseSubPath

        .StartSubPath
        .AddLineSegment -175458, -25346
        .AddLineSegment -174458, -25346
        .AddLineSegment -174458, -21737
        .AddLineSegment -175458, -21737
        .AddLineSegment -175458, -25346
        .CloseSubPath

        .StartSubPath
        .AddLineSegment -320457, -25346
        .AddLineSegment -319457, -25346
        .AddLineSegment -319457, -17824
        .AddLineSegment -320457, -17824
        .AddLineSegment -320457, -25346
        .CloseSubPath

        .StartSubPath
        .AddLineSegment -240458, -25346
        .AddLineSegment -239458, -25346
        .AddLineSegment -239458, -17824
        .AddLineSegment -240458, -17824
        .AddLineSegment -240458, -25346
        .CloseSubPath

        .StartSubPath
        .AddLineSegment -200458, -25346
        .AddLineSegment -199458, -25346
        .AddLineSegment -199458, -17824
        .AddLineSegment -200458, -17824
        .AddLineSegment -200458, -25346
        .CloseSubPath

        .StartSubPath
        .AddLineSegment -180458, -25346
        .AddLineSegment -179458, -25346
        .AddLineSegment -179458, -17824
        .AddLineSegment -180458, -17824
        .AddLineSegment -180458, -25346
        .CloseSubPath

        .StartSubPath
        .AddLineSegment -190458, -25346
        .AddLineSegment -189458, -25346
        .AddLineSegment -189458, -17824
        .AddLineSegment -190458, -17824
        .AddLineSegment -190458, -25346
        .CloseSubPath

        .StartSubPath
        .AddLineSegment -210458, -25346
        .AddLineSegment -209458, -25346
        .AddLineSegment -209458, -17824
        .AddLineSegment -210458, -17824
        .AddLineSegment -210458, -25346
        .CloseSubPath

        .StartSubPath
        .AddLineSegment -220458, -25346
        .AddLineSegment -219458, -25346
        .AddLineSegment -219458, -17824
        .AddLineSegment -220458, -17824
        .AddLineSegment -220458, -25346
        .CloseSubPath

        .StartSubPath
        .AddLineSegment -230458, -25346
        .AddLineSegment -229458, -25346
        .AddLineSegment -229458, -17824
        .AddLineSegment -230458, -17824
        .AddLineSegment -230458, -25346
        .CloseSubPath

        .StartSubPath
        .AddLineSegment -250458, -25346
        .AddLineSegment -249458, -25346
        .AddLineSegment -249458, -17824
        .AddLineSegment -250458, -17824
        .AddLineSegment -250458, -25346
        .CloseSubPath

        .StartSubPath
        .AddLineSegment -280457, -25346
        .AddLineSegment -279457, -25346
        .AddLineSegment -279457, -17824
        .AddLineSegment -280457, -17824
        .AddLineSegment -280457, -25346
        .CloseSubPath

        .StartSubPath
        .AddLineSegment -260458, -25346
        .AddLineSegment -259458, -25346
        .AddLineSegment -259458, -17824
        .AddLineSegment -260458, -17824
        .AddLineSegment -260458, -25346
        .CloseSubPath

        .StartSubPath
        .AddLineSegment -270457, -25346
        .AddLineSegment -269457, -25346
        .AddLineSegment -269457, -17824
        .AddLineSegment -270457, -17824
        .AddLineSegment -270457, -25346
        .CloseSubPath

        .StartSubPath
        .AddLineSegment -290457, -25346
        .AddLineSegment -289457, -25346
        .AddLineSegment -289457, -17824
        .AddLineSegment -290457, -17824
        .AddLineSegment -290457, -25346
        .CloseSubPath

        .StartSubPath
        .AddLineSegment -300457, -25346
        .AddLineSegment -299457, -25346
        .AddLineSegment -299457, -17824
        .AddLineSegment -300457, -17824
        .AddLineSegment -300457, -25346
        .CloseSubPath

        .StartSubPath
        .AddLineSegment -310457, -25346
        .AddLineSegment -309457, -25346
        .AddLineSegment -309457, -17824
        .AddLineSegment -310457, -17824
        .AddLineSegment -310457, -25346
        .CloseSubPath

        .StartSubPath
        .AddLineSegment -330457, -25346
        .AddLineSegment -329457, -25346
        .AddLineSegment -329457, -17824
        .AddLineSegment -330457, -17824
        .AddLineSegment -330457, -25346
        .CloseSubPath

        .StartSubPath
        .AddLineSegment -245458, -25346
        .AddLineSegment -244458, -25346
        .AddLineSegment -244458, -21737
        .AddLineSegment -245458, -21737
        .AddLineSegment -245458, -25346
        .CloseSubPath

        .StartSubPath
        .AddLineSegment -205458, -25346
        .AddLineSegment -204458, -25346
        .AddLineSegment -204458, -21737
        .AddLineSegment -205458, -21737
        .AddLineSegment -205458, -25346
        .CloseSubPath

        .StartSubPath
        .AddLineSegment -185458, -25346
        .AddLineSegment -184458, -25346
        .AddLineSegment -184458, -21737
        .AddLineSegment -185458, -21737
        .AddLineSegment -185458, -25346
        .CloseSubPath

        .StartSubPath
        .AddLineSegment -195458, -25346
        .AddLineSegment -194458, -25346
        .AddLineSegment -194458, -21737
        .AddLineSegment -195458, -21737
        .AddLineSegment -195458, -25346
        .CloseSubPath

        .StartSubPath
        .AddLineSegment -215458, -25346
        .AddLineSegment -214458, -25346
        .AddLineSegment -214458, -21737
        .AddLineSegment -215458, -21737
        .AddLineSegment -215458, -25346
        .CloseSubPath

        .StartSubPath
        .AddLineSegment -225458, -25346
        .AddLineSegment -224458, -25346
        .AddLineSegment -224458, -21737
        .AddLineSegment -225458, -21737
        .AddLineSegment -225458, -25346
        .CloseSubPath

        .StartSubPath
        .AddLineSegment -235458, -25346
        .AddLineSegment -234458, -25346
        .AddLineSegment -234458, -21737
        .AddLineSegment -235458, -21737
        .AddLineSegment -235458, -25346
        .CloseSubPath

        .StartSubPath
        .AddLineSegment -255458, -25346
        .AddLineSegment -254458, -25346
        .AddLineSegment -254458, -21737
        .AddLineSegment -255458, -21737
        .AddLineSegment -255458, -25346
        .CloseSubPath

        .StartSubPath
        .AddLineSegment -285457, -25346
        .AddLineSegment -284457, -25346
        .AddLineSegment -284457, -21737
        .AddLineSegment -285457, -21737
        .AddLineSegment -285457, -25346
        .CloseSubPath

        .StartSubPath
        .AddLineSegment -265458, -25346
        .AddLineSegment -264458, -25346
        .AddLineSegment -264458, -21737
        .AddLineSegment -265458, -21737
        .AddLineSegment -265458, -25346
        .CloseSubPath

        .StartSubPath
        .AddLineSegment -275457, -25346
        .AddLineSegment -274457, -25346
        .AddLineSegment -274457, -21737
        .AddLineSegment -275457, -21737
        .AddLineSegment -275457, -25346
        .CloseSubPath

        .StartSubPath
        .AddLineSegment -295457, -25346
        .AddLineSegment -294457, -25346
        .AddLineSegment -294457, -21737
        .AddLineSegment -295457, -21737
        .AddLineSegment -295457, -25346
        .CloseSubPath

        .StartSubPath
        .AddLineSegment -305457, -25346
        .AddLineSegment -304457, -25346
        .AddLineSegment -304457, -21737
        .AddLineSegment -305457, -21737
        .AddLineSegment -305457, -25346
        .CloseSubPath

        .StartSubPath
        .AddLineSegment -315457, -25346
        .AddLineSegment -314457, -25346
        .AddLineSegment -314457, -21737
        .AddLineSegment -315457, -21737
        .AddLineSegment -315457, -25346
        .CloseSubPath

        .StartSubPath
        .AddLineSegment -335457, -25346
        .AddLineSegment -334457, -25346
        .AddLineSegment -334457, -21737
        .AddLineSegment -335457, -21737
        .AddLineSegment -335457, -25346
        .CloseSubPath

        .StartSubPath
        .AddLineSegment -485457, -25346
        .AddLineSegment -484457, -25346
        .AddLineSegment -484457, -21737
        .AddLineSegment -485457, -21737
        .AddLineSegment -485457, -25346
        .CloseSubPath

        .StartSubPath
        .AddLineSegment -480457, -25346
        .AddLineSegment -479457, -25346
        .AddLineSegment -479457, -17824
        .AddLineSegment -480457, -17824
        .AddLineSegment -480457, -25346
        .CloseSubPath

        .StartSubPath
        .AddLineSegment -400457, -25346
        .AddLineSegment -399457, -25346
        .AddLineSegment -399457, -17824
        .AddLineSegment -400457, -17824
        .AddLineSegment -400457, -25346
        .CloseSubPath

        .StartSubPath
        .AddLineSegment -360457, -25346
        .AddLineSegment -359457, -25346
        .AddLineSegment -359457, -17824
        .AddLineSegment -360457, -17824
        .AddLineSegment -360457, -25346
        .CloseSubPath

        .StartSubPath
        .AddLineSegment -340457, -25346
        .AddLineSegment -339457, -25346
        .AddLineSegment -339457, -17824
        .AddLineSegment -340457, -17824
        .AddLineSegment -340457, -25346
        .CloseSubPath

        .StartSubPath
        .AddLineSegment -350457, -25346
        .AddLineSegment -349457, -25346
        .AddLineSegment -349457, -17824
        .AddLineSegment -350457, -17824
        .AddLineSegment -350457, -25346
        .CloseSubPath

        .StartSubPath
        .AddLineSegment -370457, -25346
        .AddLineSegment -369457, -25346
        .AddLineSegment -369457, -17824
        .AddLineSegment -370457, -17824
        .AddLineSegment -370457, -25346
        .CloseSubPath

        .StartSubPath
        .AddLineSegment -380457, -25346
        .AddLineSegment -379457, -25346
        .AddLineSegment -379457, -17824
        .AddLineSegment -380457, -17824
        .AddLineSegment -380457, -25346
        .CloseSubPath

        .StartSubPath
        .AddLineSegment -390457, -25346
        .AddLineSegment -389457, -25346
        .AddLineSegment -389457, -17824
        .AddLineSegment -390457, -17824
        .AddLineSegment -390457, -25346
        .CloseSubPath

        .StartSubPath
        .AddLineSegment -410457, -25346
        .AddLineSegment -409457, -25346
        .AddLineSegment -409457, -17824
        .AddLineSegment -410457, -17824
        .AddLineSegment -410457, -25346
        .CloseSubPath

        .StartSubPath
        .AddLineSegment -440457, -25346
        .AddLineSegment -439457, -25346
        .AddLineSegment -439457, -17824
        .AddLineSegment -440457, -17824
        .AddLineSegment -440457, -25346
        .CloseSubPath

        .StartSubPath
        .AddLineSegment -420457, -25346
        .AddLineSegment -419457, -25346
        .AddLineSegment -419457, -17824
        .AddLineSegment -420457, -17824
        .AddLineSegment -420457, -25346
        .CloseSubPath

        .StartSubPath
        .AddLineSegment -430457, -25346
        .AddLineSegment -429457, -25346
        .AddLineSegment -429457, -17824
        .AddLineSegment -430457, -17824
        .AddLineSegment -430457, -25346
        .CloseSubPath

        .StartSubPath
        .AddLineSegment -450457, -25346
        .AddLineSegment -449457, -25346
        .AddLineSegment -449457, -17824
        .AddLineSegment -450457, -17824
        .AddLineSegment -450457, -25346
        .CloseSubPath

        .StartSubPath
        .AddLineSegment -460457, -25346
        .AddLineSegment -459457, -25346
        .AddLineSegment -459457, -17824
        .AddLineSegment -460457, -17824
        .AddLineSegment -460457, -25346
        .CloseSubPath

        .StartSubPath
        .AddLineSegment -470457, -25346
        .AddLineSegment -469457, -25346
        .AddLineSegment -469457, -17824
        .AddLineSegment -470457, -17824
        .AddLineSegment -470457, -25346
        .CloseSubPath

        .StartSubPath
        .AddLineSegment -490457, -25346
        .AddLineSegment -489457, -25346
        .AddLineSegment -489457, -17824
        .AddLineSegment -490457, -17824
        .AddLineSegment -490457, -25346
        .CloseSubPath

        .StartSubPath
        .AddLineSegment -405457, -25346
        .AddLineSegment -404457, -25346
        .AddLineSegment -404457, -21737
        .AddLineSegment -405457, -21737
        .AddLineSegment -405457, -25346
        .CloseSubPath

        .StartSubPath
        .AddLineSegment -365457, -25346
        .AddLineSegment -364457, -25346
        .AddLineSegment -364457, -21737
        .AddLineSegment -365457, -21737
        .AddLineSegment -365457, -25346
        .CloseSubPath

        .StartSubPath
        .AddLineSegment -345457, -25346
        .AddLineSegment -344457, -25346
        .AddLineSegment -344457, -21737
        .AddLineSegment -345457, -21737
        .AddLineSegment -345457, -25346
        .CloseSubPath

        .StartSubPath
        .AddLineSegment -355457, -25346
        .AddLineSegment -354457, -25346
        .AddLineSegment -354457, -21737
        .AddLineSegment -355457, -21737
        .AddLineSegment -355457, -25346
        .CloseSubPath

        .StartSubPath
        .AddLineSegment -375457, -25346
        .AddLineSegment -374457, -25346
        .AddLineSegment -374457, -21737
        .AddLineSegment -375457, -21737
        .AddLineSegment -375457, -25346
        .CloseSubPath

        .StartSubPath
        .AddLineSegment -385457, -25346
        .AddLineSegment -384457, -25346
        .AddLineSegment -384457, -21737
        .AddLineSegment -385457, -21737
        .AddLineSegment -385457, -25346
        .CloseSubPath

        .StartSubPath
        .AddLineSegment -395457, -25346
        .AddLineSegment -394457, -25346
        .AddLineSegment -394457, -21737
        .AddLineSegment -395457, -21737
        .AddLineSegment -395457, -25346
        .CloseSubPath

        .StartSubPath
        .AddLineSegment -415457, -25346
        .AddLineSegment -414457, -25346
        .AddLineSegment -414457, -21737
        .AddLineSegment -415457, -21737
        .AddLineSegment -415457, -25346
        .CloseSubPath

        .StartSubPath
        .AddLineSegment -445457, -25346
        .AddLineSegment -444457, -25346
        .AddLineSegment -444457, -21737
        .AddLineSegment -445457, -21737
        .AddLineSegment -445457, -25346
        .CloseSubPath

        .StartSubPath
        .AddLineSegment -425457, -25346
        .AddLineSegment -424457, -25346
        .AddLineSegment -424457, -21737
        .AddLineSegment -425457, -21737
        .AddLineSegment -425457, -25346
        .CloseSubPath

        .StartSubPath
        .AddLineSegment -435457, -25346
        .AddLineSegment -434457, -25346
        .AddLineSegment -434457, -21737
        .AddLineSegment -435457, -21737
        .AddLineSegment -435457, -25346
        .CloseSubPath

        .StartSubPath
        .AddLineSegment -455457, -25346
        .AddLineSegment -454457, -25346
        .AddLineSegment -454457, -21737
        .AddLineSegment -455457, -21737
        .AddLineSegment -455457, -25346
        .CloseSubPath

        .StartSubPath
        .AddLineSegment -465457, -25346
        .AddLineSegment -464457, -25346
        .AddLineSegment -464457, -21737
        .AddLineSegment -465457, -21737
        .AddLineSegment -465457, -25346
        .CloseSubPath

        .StartSubPath
        .AddLineSegment -475457, -25346
        .AddLineSegment -474457, -25346
        .AddLineSegment -474457, -21737
        .AddLineSegment -475457, -21737
        .AddLineSegment -475457, -25346
        .CloseSubPath

        .StartSubPath
        .AddLineSegment -495457, -25346
        .AddLineSegment -494457, -25346
        .AddLineSegment -494457, -21737
        .AddLineSegment -495457, -21737
        .AddLineSegment -495457, -25346
        .CloseSubPath

        .StartSubPath
        .AddLineSegment -640456, -25346
        .AddLineSegment -639456, -25346
        .AddLineSegment -639456, -17824
        .AddLineSegment -640456, -17824
        .AddLineSegment -640456, -25346
        .CloseSubPath

        .StartSubPath
        .AddLineSegment -560457, -25346
        .AddLineSegment -559457, -25346
        .AddLineSegment -559457, -17824
        .AddLineSegment -560457, -17824
        .AddLineSegment -560457, -25346
        .CloseSubPath

        .StartSubPath
        .AddLineSegment -520457, -25346
        .AddLineSegment -519457, -25346
        .AddLineSegment -519457, -17824
        .AddLineSegment -520457, -17824
        .AddLineSegment -520457, -25346
        .CloseSubPath

        .StartSubPath
        .AddLineSegment -500457, -25346
        .AddLineSegment -499457, -25346
        .AddLineSegment -499457, -17824
        .AddLineSegment -500457, -17824
        .AddLineSegment -500457, -25346
        .CloseSubPath

        .StartSubPath
        .AddLineSegment -510457, -25346
        .AddLineSegment -509457, -25346
        .AddLineSegment -509457, -17824
        .AddLineSegment -510457, -17824
        .AddLineSegment -510457, -25346
        .CloseSubPath

        .StartSubPath
        .AddLineSegment -530457, -25346
        .AddLineSegment -529457, -25346
        .AddLineSegment -529457, -17824
        .AddLineSegment -530457, -17824
        .AddLineSegment -530457, -25346
        .CloseSubPath

        .StartSubPath
        .AddLineSegment -540457, -25346
        .AddLineSegment -539457, -25346
        .AddLineSegment -539457, -17824
        .AddLineSegment -540457, -17824
        .AddLineSegment -540457, -25346
        .CloseSubPath

        .StartSubPath
        .AddLineSegment -550457, -25346
        .AddLineSegment -549457, -25346
        .AddLineSegment -549457, -17824
        .AddLineSegment -550457, -17824
        .AddLineSegment -550457, -25346
        .CloseSubPath

        .StartSubPath
        .AddLineSegment -570456, -25346
        .AddLineSegment -569456, -25346
        .AddLineSegment -569456, -17824
        .AddLineSegment -570456, -17824
        .AddLineSegment -570456, -25346
        .CloseSubPath

        .StartSubPath
        .AddLineSegment -600456, -25346
        .AddLineSegment -599456, -25346
        .AddLineSegment -599456, -17824
        .AddLineSegment -600456, -17824
        .AddLineSegment -600456, -25346
        .CloseSubPath

        .StartSubPath
        .AddLineSegment -580456, -25346
        .AddLineSegment -579456, -25346
        .AddLineSegment -579456, -17824
        .AddLineSegment -580456, -17824
        .AddLineSegment -580456, -25346
        .CloseSubPath

        .StartSubPath
        .AddLineSegment -590456, -25346
        .AddLineSegment -589456, -25346
        .AddLineSegment -589456, -17824
        .AddLineSegment -590456, -17824
        .AddLineSegment -590456, -25346
        .CloseSubPath

        .StartSubPath
        .AddLineSegment -610456, -25346
        .AddLineSegment -609456, -25346
        .AddLineSegment -609456, -17824
        .AddLineSegment -610456, -17824
        .AddLineSegment -610456, -25346
        .CloseSubPath

        .StartSubPath
        .AddLineSegment -620456, -25346
        .AddLineSegment -619456, -25346
        .AddLineSegment -619456, -17824
        .AddLineSegment -620456, -17824
        .AddLineSegment -620456, -25346
        .CloseSubPath

        .StartSubPath
        .AddLineSegment -630456, -25346
        .AddLineSegment -629456, -25346
        .AddLineSegment -629456, -17824
        .AddLineSegment -630456, -17824
        .AddLineSegment -630456, -25346
        .CloseSubPath

        .StartSubPath
        .AddLineSegment -650456, -25346
        .AddLineSegment -649456, -25346
        .AddLineSegment -649456, -17824
        .AddLineSegment -650456, -17824
        .AddLineSegment -650456, -25346
        .CloseSubPath

        .StartSubPath
        .AddLineSegment -565457, -25346
        .AddLineSegment -564457, -25346
        .AddLineSegment -564457, -21737
        .AddLineSegment -565457, -21737
        .AddLineSegment -565457, -25346
        .CloseSubPath

        .StartSubPath
        .AddLineSegment -525457, -25346
        .AddLineSegment -524457, -25346
        .AddLineSegment -524457, -21737
        .AddLineSegment -525457, -21737
        .AddLineSegment -525457, -25346
        .CloseSubPath

        .StartSubPath
        .AddLineSegment -505457, -25346
        .AddLineSegment -504457, -25346
        .AddLineSegment -504457, -21737
        .AddLineSegment -505457, -21737
        .AddLineSegment -505457, -25346
        .CloseSubPath

        .StartSubPath
        .AddLineSegment -515457, -25346
        .AddLineSegment -514457, -25346
        .AddLineSegment -514457, -21737
        .AddLineSegment -515457, -21737
        .AddLineSegment -515457, -25346
        .CloseSubPath

        .StartSubPath
        .AddLineSegment -535457, -25346
        .AddLineSegment -534457, -25346
        .AddLineSegment -534457, -21737
        .AddLineSegment -535457, -21737
        .AddLineSegment -535457, -25346
        .CloseSubPath

        .StartSubPath
        .AddLineSegment -545457, -25346
        .AddLineSegment -544457, -25346
        .AddLineSegment -544457, -21737
        .AddLineSegment -545457, -21737
        .AddLineSegment -545457, -25346
        .CloseSubPath

        .StartSubPath
        .AddLineSegment -555457, -25346
        .AddLineSegment -554457, -25346
        .AddLineSegment -554457, -21737
        .AddLineSegment -555457, -21737
        .AddLineSegment -555457, -25346
        .CloseSubPath

        .StartSubPath
        .AddLineSegment -575456, -25346
        .AddLineSegment -574456, -25346
        .AddLineSegment -574456, -21737
        .AddLineSegment -575456, -21737
        .AddLineSegment -575456, -25346
        .CloseSubPath

        .StartSubPath
        .AddLineSegment -605456, -25346
        .AddLineSegment -604456, -25346
        .AddLineSegment -604456, -21737
        .AddLineSegment -605456, -21737
        .AddLineSegment -605456, -25346
        .CloseSubPath

        .StartSubPath
        .AddLineSegment -585456, -25346
        .AddLineSegment -584456, -25346
        .AddLineSegment -584456, -21737
        .AddLineSegment -585456, -21737
        .AddLineSegment -585456, -25346
        .CloseSubPath

        .StartSubPath
        .AddLineSegment -595456, -25346
        .AddLineSegment -594456, -25346
        .AddLineSegment -594456, -21737
        .AddLineSegment -595456, -21737
        .AddLineSegment -595456, -25346
        .CloseSubPath

        .StartSubPath
        .AddLineSegment -615456, -25346
        .AddLineSegment -614456, -25346
        .AddLineSegment -614456, -21737
        .AddLineSegment -615456, -21737
        .AddLineSegment -615456, -25346
        .CloseSubPath

        .StartSubPath
        .AddLineSegment -625456, -25346
        .AddLineSegment -624456, -25346
        .AddLineSegment -624456, -21737
        .AddLineSegment -625456, -21737
        .AddLineSegment -625456, -25346
        .CloseSubPath

        .StartSubPath
        .AddLineSegment -635456, -25346
        .AddLineSegment -634456, -25346
        .AddLineSegment -634456, -21737
        .AddLineSegment -635456, -21737
        .AddLineSegment -635456, -25346
        .CloseSubPath

        .StartSubPath
        .AddLineSegment -655456, -25346
        .AddLineSegment -654456, -25346
        .AddLineSegment -654456, -21737
        .AddLineSegment -655456, -21737
        .AddLineSegment -655456, -25346
        .CloseSubPath

        .StartSubPath
        .AddLineSegment -665456, -25346
        .AddLineSegment -664456, -25346
        .AddLineSegment -664456, -21737
        .AddLineSegment -665456, -21737
        .AddLineSegment -665456, -25346
        .CloseSubPath

        .StartSubPath
        .AddLineSegment -660456, -25346
        .AddLineSegment -659456, -25346
        .AddLineSegment -659456, -17824
        .AddLineSegment -660456, -17824
        .AddLineSegment -660456, -25346
        .CloseSubPath

        .StartSubPath
        .AddLineSegment 64541, -25346
        .AddLineSegment 65541, -25346
        .AddLineSegment 65541, -21737
        .AddLineSegment 64541, -21737
        .AddLineSegment 64541, -25346
        .CloseSubPath

        .StartSubPath
        .AddLineSegment 59541, -25346
        .AddLineSegment 60541, -25346
        .AddLineSegment 60541, -17824
        .AddLineSegment 59541, -17824
        .AddLineSegment 59541, -25346
        .CloseSubPath

        Set .brush = New waBrush
        .brush.SetColor RGB(114, 114, 113), 100
    End With
    MyCurves.Add path

    Set path = New waPath
    With path
        .StartSubPath
        .AddLineSegment -609419, -71852
        .AddLineSegment -610127, -57346
        .AddLineSegment -591024, -57346
        .AddLineSegment -591659, -71852
        .AddBezierSegment -591659, -71852, -595285, -72895, -605927, -73188, -609419, -71852
        .CloseSubPath

        Set .brush = New waBrush
        .brush.SetLinearGradient -620000, -65000, -601000, -65000, RGB(67, 66, 66), RGB(143, 143, 143), WrapModeTileFlipXY
    End With
    MyCurves.Add path

    Set path = New waPath
    With path
        .StartSubPath
        .AddLineSegment -10242, -71852
        .AddLineSegment -10950, -57346
        .AddLineSegment 8152, -57346
        .AddLineSegment 7517, -71852
        .AddBezierSegment 7517, -71852, 3892, -72895, -6750, -73188, -10242, -71852
        .CloseSubPath

        Set .brush = New waBrush
        .brush.SetLinearGradient -21000, -65000, -1000, -65000, RGB(67, 66, 66), RGB(143, 143, 143), WrapModeTileFlipXY
    End With
    MyCurves.Add path

    Set path = New waPath
    With path
        .StartSubPath
        .AddLineSegment -670456, 119627
        .AddLineSegment 70541, 119627
        .AddLineSegment 70541, 91572
        .AddLineSegment -670456, 91572
        .AddLineSegment -670456, 119627
        .CloseSubPath

        Set .brush = New waBrush
        .brush.SetLinearGradient -300000, 92000, -300000, 120000, RGB(254, 254, 254), RGB(233, 233, 233), WrapModeTileFlipXY
    End With
    MyCurves.Add path

    Set path = New waPath
    With path
        .StartSubPath
        .AddLineSegment -670456, 78872
        .AddLineSegment 70541, 78872
        .AddLineSegment 70541, 91837
        .AddLineSegment -670456, 91837
        .AddLineSegment -670456, 78872
        .CloseSubPath

        Set .brush = New waBrush
        .brush.SetColor RGB(254, 254, 254), 100
    End With
    MyCurves.Add path

    Set path = New waPath
    With path
        .StartSubPath
        .AddLineSegment -647881, 81139
        .AddBezierSegment -647881, 81139, -645660, 81139, -643859, 82939, -643859, 85160
        .AddBezierSegment -643859, 85160, -643859, 87381, -645660, 89182, -647881, 89182
        .AddBezierSegment -647881, 89182, -650102, 89182, -651902, 87381, -651902, 85160
        .AddBezierSegment -651902, 85160, -651902, 82939, -650102, 81139, -647881, 81139
        .CloseSubPath

        Set .brush = New waBrush
        .brush.SetLinearGradient -648000, 85000, -644000, 85000, RGB(217, 218, 218), RGB(242, 239, 238), WrapModeTileFlipXY
    End With
    MyCurves.Add path

    Set path = New waPath
    With path
        .StartSubPath
        .AddLineSegment -648810, 86273
        .AddLineSegment -648380, 86522
        .AddLineSegment -647790, 85500
        .AddLineSegment -646768, 86090
        .AddLineSegment -646519, 85659
        .AddLineSegment -647541, 85069
        .AddLineSegment -646951, 84047
        .AddLineSegment -647381, 83799
        .AddLineSegment -647972, 84821
        .AddLineSegment -648994, 84231
        .AddLineSegment -649242, 84661
        .AddLineSegment -648220, 85251
        .AddLineSegment -648810, 86273
        .CloseSubPath

        Set .brush = New waBrush
        .brush.SetColor RGB(197, 198, 198), 100
    End With
    MyCurves.Add path

    Set path = New waPath
    With path
        .StartSubPath
        .AddLineSegment 48491, 82087
        .AddBezierSegment 48491, 82087, 50188, 80654, 52726, 80869, 54158, 82566
        .AddBezierSegment 54158, 82566, 55591, 84263, 55376, 86801, 53679, 88234
        .AddBezierSegment 53679, 88234, 51982, 89666, 49444, 89452, 48012, 87754
        .AddBezierSegment 48012, 87754, 46579, 86057, 46794, 83520, 48491, 82087
        .CloseSubPath

        Set .brush = New waBrush
        .brush.SetLinearGradient 51000, 85000, 54000, 83000, RGB(217, 218, 218), RGB(242, 239, 238), WrapModeTileFlipXY
    End With
    MyCurves.Add path

    Set path = New waPath
    With path
        .StartSubPath
        .AddLineSegment 51093, 86611
        .AddLineSegment 51582, 86523
        .AddLineSegment 51373, 85361
        .AddLineSegment 52535, 85153
        .AddLineSegment 52447, 84663
        .AddLineSegment 51286, 84872
        .AddLineSegment 51077, 83710
        .AddLineSegment 50588, 83798
        .AddLineSegment 50797, 84960
        .AddLineSegment 49635, 85168
        .AddLineSegment 49722, 85657
        .AddLineSegment 50884, 85449
        .AddLineSegment 51093, 86611
        .CloseSubPath

        Set .brush = New waBrush
        .brush.SetColor RGB(197, 198, 198), 100
    End With
    MyCurves.Add path

    Set path = New waPath
    With path
        .StartSubPath
        .AddLineSegment -670456, -17824
        .AddLineSegment 70541, -17824
        .AddLineSegment 70541, 0
        .AddLineSegment -670456, 0
        .AddLineSegment -670456, -17824
        .CloseSubPath

        Set .brush = New waBrush
        .brush.SetLinearGradient -300000, -18000, -300000, 0, RGB(164, 185, 195), RGB(148, 148, 148), WrapModeTileFlipXY
    End With
    MyCurves.Add path

    Set path = New waPath
    With path
        .StartSubPath
        .AddLineSegment -628379, 27225
        .AddLineSegment 70541, 27225
        .AddLineSegment 70541, 22145
        .AddLineSegment -628379, 22145
        .AddLineSegment -628379, 27225
        .CloseSubPath

        Set .brush = New waBrush
        .brush.SetLinearGradient -279000, 25000, -279000, 20000, RGB(43, 42, 41), RGB(194, 194, 194), WrapModeTileFlipXY
    End With
    MyCurves.Add path

    Set path = New waPath
    With path
        .StartSubPath
        .AddLineSegment 70541, 66171
        .AddLineSegment -670456, 66171
        .AddLineSegment -670456, 27225
        .AddLineSegment 70541, 27225
        .AddLineSegment 70541, 66171
        .CloseSubPath

        Set .brush = New waBrush
        .brush.SetColor RGB(67, 66, 66), 100
    End With
    MyCurves.Add path

    Set path = New waPath
    With path
        .StartSubPath
        .AddLineSegment -670456, 46250
        .AddLineSegment 70541, 46250
        .AddLineSegment 70541, 36538
        .AddLineSegment -670456, 36538
        .AddLineSegment -670456, 46250
        .CloseSubPath

        Set .brush = New waBrush
        .brush.SetColor RGB(43, 42, 41), 100
    End With
    MyCurves.Add path

    Set path = New waPath
    With path
        .StartSubPath
        .AddLineSegment -627158, 32622
        .AddLineSegment -579745, 32622
        .AddLineSegment -579745, 27225
        .AddLineSegment -627158, 27225
        .AddLineSegment -627158, 32622
        .CloseSubPath

        .StartSubPath
        .AddLineSegment -8192, 32622
        .AddLineSegment 39222, 32622
        .AddLineSegment 39222, 27225
        .AddLineSegment -8192, 27225
        .AddLineSegment -8192, 32622
        .CloseSubPath

        .StartSubPath
        .AddLineSegment -112014, 32622
        .AddLineSegment -64601, 32622
        .AddLineSegment -64601, 27225
        .AddLineSegment -112014, 27225
        .AddLineSegment -112014, 32622
        .CloseSubPath

        .StartSubPath
        .AddLineSegment -189907, 32622
        .AddLineSegment -142494, 32622
        .AddLineSegment -142494, 27225
        .AddLineSegment -189907, 27225
        .AddLineSegment -189907, 32622
        .CloseSubPath

        .StartSubPath
        .AddLineSegment -382630, 32622
        .AddLineSegment -335217, 32622
        .AddLineSegment -335217, 27225
        .AddLineSegment -382630, 27225
        .AddLineSegment -382630, 32622
        .CloseSubPath

        Set .brush = New waBrush
        .brush.SetColor RGB(254, 254, 254), 100
    End With
    MyCurves.Add path

    Set path = New waPath
    With path
        .StartSubPath
        .AddLineSegment -314034, 0
        .AddLineSegment -285883, 0
        .AddLineSegment -285883, 6430
        .AddLineSegment -289587, 6430
        .AddLineSegment -289587, 22145
        .AddLineSegment -310330, 22145
        .AddLineSegment -310330, 6430
        .AddLineSegment -314034, 6430
        .AddLineSegment -314034, 0
        .CloseSubPath

        Set .brush = New waBrush
        .brush.SetColor RGB(43, 42, 41), 100
    End With
    MyCurves.Add path

    Set path = New waPath
    With path
        .StartSubPath
        .AddLineSegment -625000, 22145
        .AddLineSegment -599971, 22145
        .AddLineSegment -599971, 2567
        .AddBezierSegment -599971, 2567, -599971, 2014, -600423, 1562, -600976, 1562
        .AddLineSegment -605738, 1562
        .AddBezierSegment -605738, 1562, -606291, 1562, -606744, 2014, -606744, 2567
        .AddLineSegment -606744, 9711
        .AddLineSegment -618227, 9711
        .AddLineSegment -618227, 2567
        .AddBezierSegment -618227, 2567, -618227, 2014, -618679, 1562, -619232, 1562
        .AddLineSegment -623995, 1562
        .AddBezierSegment -623995, 1562, -624548, 1562, -625000, 2014, -625000, 2567
        .AddBezierSegment -625000, 2567, -625000, 9093, -625000, 15619, -625000, 22145
        .CloseSubPath

        Set .brush = New waBrush
        .brush.SetColor RGB(43, 42, 41), 100
    End With
    MyCurves.Add path

    Set path = New waPath
    With path
        .StartSubPath
        .AddLineSegment -615528, 0
        .AddLineSegment -609389, 0
        .AddBezierSegment -609389, 0, -608109, 0, -607061, 1048, -607061, 2328
        .AddLineSegment -607061, 10716
        .AddBezierSegment -607061, 10716, -607061, 11997, -608109, 13045, -609389, 13045
        .AddLineSegment -615528, 13045
        .AddBezierSegment -615528, 13045, -616808, 13045, -617856, 11997, -617856, 10716
        .AddLineSegment -617856, 2328
        .AddBezierSegment -617856, 2328, -617856, 1048, -616808, 0, -615528, 0
        .CloseSubPath

        Set .brush = New waBrush
        .brush.SetLinearGradient -612000, -7000, -612000, 7000, RGB(43, 42, 41), RGB(167, 165, 165), WrapModeTileFlipXY
    End With
    MyCurves.Add path

    Set path = New waPath
    With path
        .StartSubPath
        .AddLineSegment 0, 22145
        .AddLineSegment 25029, 22145
        .AddLineSegment 25029, 2567
        .AddBezierSegment 25029, 2567, 25029, 2014, 24577, 1562, 24024, 1562
        .AddLineSegment 19262, 1562
        .AddBezierSegment 19262, 1562, 18709, 1562, 18256, 2014, 18256, 2567
        .AddLineSegment 18256, 9711
        .AddLineSegment 6773, 9711
        .AddLineSegment 6773, 2567
        .AddBezierSegment 6773, 2567, 6773, 2014, 6321, 1562, 5768, 1562
        .AddLineSegment 1005, 1562
        .AddBezierSegment 1005, 1562, 452, 1562, 0, 2014, 0, 2567
        .AddBezierSegment 0, 2567, 0, 9093, 0, 15619, 0, 22145
        .CloseSubPath

        Set .brush = New waBrush
        .brush.SetColor RGB(43, 42, 41), 100
    End With
    MyCurves.Add path

    Set path = New waPath
    With path
        .StartSubPath
        .AddLineSegment 9472, 0
        .AddLineSegment 15611, 0
        .AddBezierSegment 15611, 0, 16891, 0, 17939, 1048, 17939, 2328
        .AddLineSegment 17939, 10716
        .AddBezierSegment 17939, 10716, 17939, 11997, 16891, 13045, 15611, 13045
        .AddLineSegment 9472, 13045
        .AddBezierSegment 9472, 13045, 8192, 13045, 7144, 11997, 7144, 10716
        .AddLineSegment 7144, 2328
        .AddBezierSegment 7144, 2328, 7144, 1048, 8192, 0, 9472, 0
        .CloseSubPath

        Set .brush = New waBrush
        .brush.SetLinearGradient 13000, -7000, 13000, 7000, RGB(43, 42, 41), RGB(167, 165, 165), WrapModeTileFlipXY
    End With
    MyCurves.Add path

    Set path = New waPath
    With path
        .StartSubPath
        .AddLineSegment -165020, 118871
        .AddLineSegment 20612, 118871
        .AddBezierSegment 20612, 118871, 29976, 118871, 37588, 116207, 37588, 112929
        .AddLineSegment 37588, 97513
        .AddBezierSegment 37588, 97513, 37588, 94236, 29976, 91572, 20612, 91572
        .AddLineSegment -206719, 91572
        .AddBezierSegment -206719, 91572, -216701, 91572, -225121, 94746, -223523, 98279
        .AddBezierSegment -223523, 98279, -220642, 104649, -213013, 110667, -198191, 114700
        .AddBezierSegment -198191, 114700, -188951, 117214, -176810, 118871, -165020, 118871
        .CloseSubPath

        Set .brush = New waBrush
        .brush.SetColor RGB(235, 236, 236), 100
    End With
    MyCurves.Add path

    Set path = New waPath
    With path
        .StartSubPath
        .AddLineSegment -152743, 113971
        .AddLineSegment 9605, 113971
        .AddBezierSegment 9605, 113971, 11242, 113971, 12582, 113502, 12582, 112929
        .AddLineSegment 12582, 97513
        .AddBezierSegment 12582, 97513, 12582, 96941, 11242, 96472, 9605, 96472
        .AddLineSegment -194442, 96472
        .AddBezierSegment -194442, 96472, -196079, 96472, -197639, 97026, -197418, 97513
        .AddBezierSegment -197418, 97513, -191280, 111085, -164857, 113971, -152743, 113971
        .CloseSubPath

        Set .brush = New waBrush
        .brush.SetColor RGB(43, 42, 41), 100
    End With
    MyCurves.Add path

    Set path = New waPath
    With path
        .StartSubPath
        .AddLineSegment -21964, 99119
        .AddLineSegment -11366, 99119
        .AddBezierSegment -11366, 99119, -10937, 99119, -10586, 99242, -10586, 99392
        .AddLineSegment -10586, 102429
        .AddLineSegment -16665, 104369
        .AddLineSegment -22743, 102429
        .AddLineSegment -22743, 99392
        .AddBezierSegment -22743, 99392, -22743, 99242, -22392, 99119, -21964, 99119
        .CloseSubPath

        Set .brush = New waBrush
        .brush.SetColor RGB(162, 217, 247), 100
    End With
    MyCurves.Add path

    Set path = New waPath
    With path
        .StartSubPath
        .AddLineSegment -16565, 100713
        .AddLineSegment -19713, 102201
        .AddLineSegment -13416, 102201
        .AddLineSegment -16565, 100713
        .CloseSubPath

        Set .brush = New waBrush
        .brush.SetColor RGB(43, 42, 41), 100
    End With
    MyCurves.Add path

    Set path = New waPath
    With path
        .StartSubPath
        .AddLineSegment -34074, 107067
        .AddLineSegment -34074, 103358
        .AddBezierSegment -34074, 103358, -34074, 103207, -33723, 103085, -33295, 103085
        .AddLineSegment -24616, 103085
        .AddLineSegment -19074, 105212
        .AddLineSegment -24616, 107340
        .AddLineSegment -33295, 107340
        .AddBezierSegment -33295, 107340, -33723, 107340, -34074, 107217, -34074, 107067
        .CloseSubPath

        Set .brush = New waBrush
        .brush.SetColor RGB(162, 217, 247), 100
    End With
    MyCurves.Add path

    Set path = New waPath
    With path
        .StartSubPath
        .AddLineSegment -29520, 105177
        .AddLineSegment -25268, 106279
        .AddLineSegment -25268, 104075
        .AddLineSegment -29520, 105177
        .CloseSubPath

        Set .brush = New waBrush
        .brush.SetColor RGB(43, 42, 41), 100
    End With
    MyCurves.Add path

    Set path = New waPath
    With path
        .StartSubPath
        .AddLineSegment -21964, 111305
        .AddLineSegment -11366, 111305
        .AddBezierSegment -11366, 111305, -10937, 111305, -10586, 111183, -10586, 111032
        .AddLineSegment -10586, 107995
        .AddLineSegment -16665, 106055
        .AddLineSegment -22743, 107995
        .AddLineSegment -22743, 111032
        .AddBezierSegment -22743, 111032, -22743, 111182, -22392, 111305, -21964, 111305
        .CloseSubPath

        Set .brush = New waBrush
        .brush.SetColor RGB(162, 217, 247), 100
    End With
    MyCurves.Add path

    Set path = New waPath
    With path
        .StartSubPath
        .AddLineSegment -16565, 109711
        .AddLineSegment -19713, 108223
        .AddLineSegment -13416, 108223
        .AddLineSegment -16565, 109711
        .CloseSubPath

        Set .brush = New waBrush
        .brush.SetColor RGB(43, 42, 41), 100
    End With
    MyCurves.Add path

    Set path = New waPath
    With path
        .StartSubPath
        .AddLineSegment 745, 107067
        .AddLineSegment 745, 103358
        .AddBezierSegment 745, 103358, 745, 103207, 394, 103085, -35, 103085
        .AddLineSegment -8713, 103085
        .AddLineSegment -14255, 105212
        .AddLineSegment -8713, 107340
        .AddLineSegment -35, 107340
        .AddBezierSegment -35, 107340, 394, 107340, 745, 107217, 745, 107067
        .CloseSubPath

        Set .brush = New waBrush
        .brush.SetColor RGB(162, 217, 247), 100
    End With
    MyCurves.Add path

    Set path = New waPath
    With path
        .StartSubPath
        .AddLineSegment -3809, 105177
        .AddLineSegment -8062, 106279
        .AddLineSegment -8062, 104075
        .AddLineSegment -3809, 105177
        .CloseSubPath

        Set .brush = New waBrush
        .brush.SetColor RGB(43, 42, 41), 100
    End With
    MyCurves.Add path

    Set path = New waPath
    With path
        .StartSubPath
        .AddLineSegment -49527, 100274
        .AddBezierSegment -49527, 100274, -45937, 100274, -43027, 101293, -43027, 102549
        .AddBezierSegment -43027, 102549, -43027, 103806, -45937, 104824, -49527, 104824
        .AddBezierSegment -49527, 104824, -53117, 104824, -56027, 103806, -56027, 102549
        .AddBezierSegment -56027, 102549, -56027, 101293, -53117, 100274, -49527, 100274
        .CloseSubPath

        Set .brush = New waBrush
        .brush.SetColor RGB(235, 236, 236), 100
    End With
    MyCurves.Add path

    Set path = New waPath
    With path
        .StartSubPath
        .AddLineSegment -69000, 100274
        .AddBezierSegment -69000, 100274, -65410, 100274, -62500, 101293, -62500, 102549
        .AddBezierSegment -62500, 102549, -62500, 103806, -65410, 104824, -69000, 104824
        .AddBezierSegment -69000, 104824, -72590, 104824, -75500, 103806, -75500, 102549
        .AddBezierSegment -75500, 102549, -75500, 101293, -72590, 100274, -69000, 100274
        .CloseSubPath

        Set .brush = New waBrush
        .brush.SetColor RGB(235, 236, 236), 100
    End With
    MyCurves.Add path

    Set path = New waPath
    With path
        .StartSubPath
        .AddLineSegment -88474, 100274
        .AddBezierSegment -88474, 100274, -84884, 100274, -81974, 101293, -81974, 102549
        .AddBezierSegment -81974, 102549, -81974, 103806, -84884, 104824, -88474, 104824
        .AddBezierSegment -88474, 104824, -92064, 104824, -94974, 103806, -94974, 102549
        .AddBezierSegment -94974, 102549, -94974, 101293, -92064, 100274, -88474, 100274
        .CloseSubPath

        Set .brush = New waBrush
        .brush.SetColor RGB(235, 236, 236), 100
    End With
    MyCurves.Add path

    Set path = New waPath
    With path
        .StartSubPath
        .AddLineSegment -108159, 100274
        .AddBezierSegment -108159, 100274, -104569, 100274, -101659, 101293, -101659, 102549
        .AddBezierSegment -101659, 102549, -101659, 103806, -104569, 104824, -108159, 104824
        .AddBezierSegment -108159, 104824, -111749, 104824, -114659, 103806, -114659, 102549
        .AddBezierSegment -114659, 102549, -114659, 101293, -111749, 100274, -108159, 100274
        .CloseSubPath

        Set .brush = New waBrush
        .brush.SetColor RGB(235, 236, 236), 100
    End With
    MyCurves.Add path

    Set path = New waPath
    With path
        .StartSubPath
        .AddLineSegment -127632, 100274
        .AddBezierSegment -127632, 100274, -124042, 100274, -121132, 101293, -121132, 102549
        .AddBezierSegment -121132, 102549, -121132, 103806, -124042, 104824, -127632, 104824
        .AddBezierSegment -127632, 104824, -131222, 104824, -134132, 103806, -134132, 102549
        .AddBezierSegment -134132, 102549, -134132, 101293, -131222, 100274, -127632, 100274
        .CloseSubPath

        Set .brush = New waBrush
        .brush.SetColor RGB(235, 236, 236), 100
    End With
    MyCurves.Add path

    Set path = New waPath
    With path
        .StartSubPath
        .AddLineSegment -146682, 100274
        .AddBezierSegment -146682, 100274, -143092, 100274, -140182, 101293, -140182, 102549
        .AddBezierSegment -140182, 102549, -140182, 103806, -143092, 104824, -146682, 104824
        .AddBezierSegment -146682, 104824, -150272, 104824, -153182, 103806, -153182, 102549
        .AddBezierSegment -153182, 102549, -153182, 101293, -150272, 100274, -146682, 100274
        .CloseSubPath

        Set .brush = New waBrush
        .brush.SetColor RGB(235, 236, 236), 100
    End With
    MyCurves.Add path

    Set path = New waPath
    With path
        .StartSubPath
        .AddLineSegment -161269, 107587
        .AddBezierSegment -161269, 107587, -160164, 107587, -159269, 107900, -159269, 108287
        .AddBezierSegment -159269, 108287, -159269, 108673, -160164, 108987, -161269, 108987
        .AddBezierSegment -161269, 108987, -162373, 108987, -163269, 108673, -163269, 108287
        .AddBezierSegment -163269, 108287, -163269, 107900, -162373, 107587, -161269, 107587
        .CloseSubPath

        Set .brush = New waBrush
        .brush.SetColor RGB(219, 226, 133), 100
    End With
    MyCurves.Add path

    Set path = New waPath
    With path
        .StartSubPath
        .AddLineSegment -147105, 107587
        .AddBezierSegment -147105, 107587, -146001, 107587, -145105, 107900, -145105, 108287
        .AddBezierSegment -145105, 108287, -145105, 108673, -146001, 108987, -147105, 108987
        .AddBezierSegment -147105, 108987, -148210, 108987, -149105, 108673, -149105, 108287
        .AddBezierSegment -149105, 108287, -149105, 107900, -148210, 107587, -147105, 107587
        .CloseSubPath

        Set .brush = New waBrush
        .brush.SetColor RGB(219, 226, 133), 100
    End With
    MyCurves.Add path


    InitStandardPlotter600_2DModel_2 MyCurves
    StandardPlotter600_2DModel_Initialized = True
End Sub

' -----------------------------------------------------------------------------
' InitStandardPlotter600_2DModel_2 (Sub), source lines 2732-3976
' -----------------------------------------------------------------------------
Private Sub InitStandardPlotter600_2DModel_2(MyCurves As Collection)
    Dim path As waPath
    Set path = New waPath
    With path
        .StartSubPath
        .AddLineSegment -644022, 104791
        .AddLineSegment -639475, 104966
        .AddBezierSegment -639475, 104966, -639200, 104374, -638652, 103942, -637819, 103663
        .AddBezierSegment -637819, 103663, -636986, 103385, -635863, 103247, -634449, 103247
        .AddBezierSegment -634449, 103247, -632955, 103247, -631827, 103370, -631069, 103617
        .AddBezierSegment -631069, 103617, -630311, 103864, -629929, 104153, -629929, 104483
        .AddBezierSegment -629929, 104483, -629929, 104697, -630091, 104877, -630408, 105027
        .AddBezierSegment -630408, 105027, -630725, 105176, -631284, 105305, -632079, 105414
        .AddBezierSegment -632079, 105414, -632622, 105490, -633858, 105619, -635793, 105808
        .AddBezierSegment -635793, 105808, -638281, 106050, -640028, 106347, -641033, 106701
        .AddBezierSegment -641033, 106701, -642447, 107197, -643151, 107801, -643151, 108514
        .AddBezierSegment -643151, 108514, -643151, 108973, -642818, 109403, -642157, 109803
        .AddBezierSegment -642157, 109803, -641490, 110203, -640533, 110506, -639281, 110715
        .AddBezierSegment -639281, 110715, -638029, 110924, -636524, 111029, -634750, 111029
        .AddBezierSegment -634750, 111029, -631864, 111029, -629687, 110780, -628231, 110280
        .AddBezierSegment -628231, 110280, -626769, 109780, -626006, 109115, -625930, 108280
        .AddLineSegment -630606, 108217
        .AddBezierSegment -630606, 108217, -630805, 108682, -631230, 109014, -631880, 109217
        .AddBezierSegment -631880, 109217, -632530, 109420, -633509, 109523, -634809, 109523
        .AddBezierSegment -634809, 109523, -636153, 109523, -637207, 109414, -637964, 109196
        .AddBezierSegment -637964, 109196, -638453, 109056, -638701, 108870, -638701, 108636
        .AddBezierSegment -638701, 108636, -638701, 108422, -638470, 108240, -638007, 108090
        .AddBezierSegment -638007, 108090, -637422, 107895, -635997, 107694, -633734, 107485
        .AddBezierSegment -633734, 107485, -631472, 107278, -629800, 107061, -628720, 106839
        .AddBezierSegment -628720, 106839, -627634, 106615, -626785, 106312, -626172, 105925
        .AddBezierSegment -626172, 105925, -625559, 105538, -625253, 105061, -625253, 104494
        .AddBezierSegment -625253, 104494, -625253, 103979, -625619, 103496, -626355, 103046
        .AddBezierSegment -626355, 103046, -627091, 102596, -628129, 102262, -629472, 102044
        .AddBezierSegment -629472, 102044, -630816, 101824, -632493, 101716, -634498, 101716
        .AddBezierSegment -634498, 101716, -637416, 101716, -639657, 101977, -641221, 102502
        .AddBezierSegment -641221, 102502, -642785, 103025, -643715, 103789, -644022, 104791
        .CloseSubPath

        .StartSubPath
        .AddLineSegment -616095, 101879
        .AddLineSegment -616095, 109347
        .AddLineSegment -622931, 109347
        .AddLineSegment -622931, 110866
        .AddLineSegment -604614, 110866
        .AddLineSegment -604614, 109347
        .AddLineSegment -611419, 109347
        .AddLineSegment -611419, 101879
        .AddLineSegment -616095, 101879
        .CloseSubPath

        .StartSubPath
        .AddLineSegment -583180, 101879
        .AddLineSegment -588216, 101879
        .AddLineSegment -590210, 103925
        .AddLineSegment -599471, 103925
        .AddLineSegment -601389, 101879
        .AddLineSegment -606334, 101879
        .AddLineSegment -597369, 110866
        .AddLineSegment -592430, 110866
        .AddLineSegment -583180, 101879
        .CloseSubPath

        .StartSubPath
        .AddLineSegment -591726, 105444
        .AddLineSegment -594935, 108770
        .AddLineSegment -598052, 105444
        .AddLineSegment -591726, 105444
        .CloseSubPath

        .StartSubPath
        .AddLineSegment -580692, 101879
        .AddLineSegment -580692, 110866
        .AddLineSegment -576166, 110866
        .AddLineSegment -566664, 104843
        .AddLineSegment -566664, 110866
        .AddLineSegment -562343, 110866
        .AddLineSegment -562343, 101879
        .AddLineSegment -567013, 101879
        .AddLineSegment -576370, 107787
        .AddLineSegment -576370, 101879
        .AddLineSegment -580692, 101879
        .CloseSubPath

        .StartSubPath
        .AddLineSegment -557403, 110866
        .AddLineSegment -548901, 110866
        .AddBezierSegment -548901, 110866, -546982, 110866, -545520, 110809, -544515, 110694
        .AddBezierSegment -544515, 110694, -543166, 110540, -542005, 110263, -541043, 109866
        .AddBezierSegment -541043, 109866, -540075, 109471, -539345, 108985, -538839, 108410
        .AddBezierSegment -538839, 108410, -538339, 107837, -538087, 107127, -538087, 106287
        .AddBezierSegment -538087, 106287, -538087, 105546, -538323, 104910, -538796, 104374
        .AddBezierSegment -538796, 104374, -539371, 103720, -540194, 103191, -541263, 102787
        .AddBezierSegment -541263, 102787, -542069, 102479, -543160, 102241, -544531, 102069
        .AddBezierSegment -544531, 102069, -545558, 101942, -546928, 101879, -548648, 101879
        .AddLineSegment -557403, 101879
        .AddLineSegment -557403, 110866
        .CloseSubPath

        .StartSubPath
        .AddLineSegment -552727, 109347
        .AddLineSegment -552727, 103398
        .AddLineSegment -549261, 103398
        .AddBezierSegment -549261, 103398, -547966, 103398, -547030, 103427, -546455, 103483
        .AddBezierSegment -546455, 103483, -545697, 103556, -545074, 103682, -544579, 103858
        .AddBezierSegment -544579, 103858, -544080, 104034, -543676, 104322, -543365, 104724
        .AddBezierSegment -543365, 104724, -543048, 105126, -542892, 105676, -542892, 106370
        .AddBezierSegment -542892, 106370, -542892, 107065, -543048, 107596, -543365, 107969
        .AddBezierSegment -543365, 107969, -543676, 108341, -544117, 108629, -544682, 108839
        .AddBezierSegment -544682, 108839, -545246, 109046, -545961, 109188, -546831, 109261
        .AddBezierSegment -546831, 109261, -547476, 109318, -548750, 109347, -550642, 109347
        .AddLineSegment -552727, 109347
        .CloseSubPath

        .StartSubPath
        .AddLineSegment -513315, 101879
        .AddLineSegment -518351, 101879
        .AddLineSegment -520345, 103925
        .AddLineSegment -529606, 103925
        .AddLineSegment -531524, 101879
        .AddLineSegment -536469, 101879
        .AddLineSegment -527504, 110866
        .AddLineSegment -522565, 110866
        .AddLineSegment -513315, 101879
        .CloseSubPath

        .StartSubPath
        .AddLineSegment -521861, 105444
        .AddLineSegment -525069, 108770
        .AddLineSegment -528187, 105444
        .AddLineSegment -521861, 105444
        .CloseSubPath

        .StartSubPath
        .AddLineSegment -510826, 101879
        .AddLineSegment -510826, 110866
        .AddLineSegment -501007, 110866
        .AddBezierSegment -501007, 110866, -498540, 110866, -496745, 110784, -495627, 110623
        .AddBezierSegment -495627, 110623, -494504, 110462, -493611, 110175, -492940, 109763
        .AddBezierSegment -492940, 109763, -492268, 109349, -491929, 108879, -491929, 108347
        .AddBezierSegment -491929, 108347, -491929, 107671, -492435, 107115, -493450, 106676
        .AddBezierSegment -493450, 106676, -494466, 106236, -495982, 105960, -498002, 105845
        .AddBezierSegment -498002, 105845, -496997, 105617, -496170, 105366, -495520, 105092
        .AddBezierSegment -495520, 105092, -494864, 104818, -493982, 104331, -492875, 103632
        .AddLineSegment -490059, 101879
        .AddLineSegment -495632, 101879
        .AddLineSegment -498975, 103841
        .AddBezierSegment -498975, 103841, -500174, 104538, -500996, 104977, -501442, 105159
        .AddBezierSegment -501442, 105159, -501883, 105339, -502351, 105464, -502845, 105531
        .AddBezierSegment -502845, 105531, -503345, 105598, -504130, 105632, -505204, 105632
        .AddLineSegment -506151, 105632
        .AddLineSegment -506151, 101879
        .AddLineSegment -510826, 101879
        .CloseSubPath

        .StartSubPath
        .AddLineSegment -506151, 107063
        .AddLineSegment -502700, 107063
        .AddBezierSegment -502700, 107063, -500464, 107063, -499072, 107100, -498513, 107174
        .AddBezierSegment -498513, 107174, -497960, 107247, -497519, 107374, -497207, 107554
        .AddBezierSegment -497207, 107554, -496890, 107734, -496734, 107960, -496734, 108230
        .AddBezierSegment -496734, 108230, -496734, 108533, -496944, 108776, -497358, 108964
        .AddBezierSegment -497358, 108964, -497772, 109150, -498357, 109268, -499110, 109316
        .AddBezierSegment -499110, 109316, -499491, 109337, -500625, 109347, -502512, 109347
        .AddLineSegment -506151, 109347
        .AddLineSegment -506151, 107063
        .CloseSubPath

        .StartSubPath
        .AddLineSegment -487538, 110866
        .AddLineSegment -479036, 110866
        .AddBezierSegment -479036, 110866, -477116, 110866, -475655, 110809, -474650, 110694
        .AddBezierSegment -474650, 110694, -473301, 110540, -472140, 110263, -471178, 109866
        .AddBezierSegment -471178, 109866, -470210, 109471, -469479, 108985, -468974, 108410
        .AddBezierSegment -468974, 108410, -468474, 107837, -468222, 107127, -468222, 106287
        .AddBezierSegment -468222, 106287, -468222, 105546, -468458, 104910, -468931, 104374
        .AddBezierSegment -468931, 104374, -469506, 103720, -470328, 103191, -471398, 102787
        .AddBezierSegment -471398, 102787, -472204, 102479, -473295, 102241, -474666, 102069
        .AddBezierSegment -474666, 102069, -475692, 101942, -477063, 101879, -478783, 101879
        .AddLineSegment -487538, 101879
        .AddLineSegment -487538, 110866
        .CloseSubPath

        .StartSubPath
        .AddLineSegment -482862, 109347
        .AddLineSegment -482862, 103398
        .AddLineSegment -479395, 103398
        .AddBezierSegment -479395, 103398, -478100, 103398, -477165, 103427, -476590, 103483
        .AddBezierSegment -476590, 103483, -475832, 103556, -475209, 103682, -474714, 103858
        .AddBezierSegment -474714, 103858, -474214, 104034, -473811, 104322, -473500, 104724
        .AddBezierSegment -473500, 104724, -473182, 105126, -473027, 105676, -473027, 106370
        .AddBezierSegment -473027, 106370, -473027, 107065, -473182, 107596, -473500, 107969
        .AddBezierSegment -473500, 107969, -473811, 108341, -474252, 108629, -474816, 108839
        .AddBezierSegment -474816, 108839, -475380, 109046, -476096, 109188, -476966, 109261
        .AddBezierSegment -476966, 109261, -477611, 109318, -478885, 109347, -480777, 109347
        .AddLineSegment -482862, 109347
        .CloseSubPath

        .StartSubPath
        .AddLineSegment -455290, 101879
        .AddLineSegment -455290, 110866
        .AddLineSegment -447814, 110866
        .AddBezierSegment -447814, 110866, -444981, 110866, -443138, 110822, -442278, 110732
        .AddBezierSegment -442278, 110732, -440956, 110596, -439849, 110303, -438957, 109851
        .AddBezierSegment -438957, 109851, -438064, 109399, -437618, 108816, -437618, 108100
        .AddBezierSegment -437618, 108100, -437618, 107548, -437876, 107086, -438387, 106709
        .AddBezierSegment -438387, 106709, -438903, 106333, -439553, 106038, -440343, 105822
        .AddBezierSegment -440343, 105822, -441133, 105609, -441939, 105466, -442756, 105397
        .AddBezierSegment -442756, 105397, -443869, 105312, -445476, 105268, -447583, 105268
        .AddLineSegment -450614, 105268
        .AddLineSegment -450614, 101879
        .AddLineSegment -455290, 101879
        .CloseSubPath

        .StartSubPath
        .AddLineSegment -450614, 109347
        .AddLineSegment -450614, 106787
        .AddLineSegment -448072, 106787
        .AddBezierSegment -448072, 106787, -446239, 106787, -445019, 106835, -444401, 106929
        .AddBezierSegment -444401, 106929, -443783, 107023, -443299, 107169, -442950, 107372
        .AddBezierSegment -442950, 107372, -442600, 107573, -442423, 107805, -442423, 108073
        .AddBezierSegment -442423, 108073, -442423, 108401, -442670, 108671, -443159, 108885
        .AddBezierSegment -443159, 108885, -443654, 109098, -444272, 109232, -445030, 109286
        .AddBezierSegment -445030, 109286, -445584, 109326, -446696, 109347, -448373, 109347
        .AddLineSegment -450614, 109347
        .CloseSubPath

        .StartSubPath
        .AddLineSegment -433652, 101879
        .AddLineSegment -433652, 110790
        .AddLineSegment -428976, 110790
        .AddLineSegment -428976, 103398
        .AddLineSegment -417399, 103398
        .AddLineSegment -417399, 101879
        .AddLineSegment -433652, 101879
        .CloseSubPath

        .StartSubPath
        .AddLineSegment -415050, 106318
        .AddBezierSegment -415050, 106318, -415050, 107234, -414695, 108004, -413991, 108627
        .AddBezierSegment -413991, 108627, -413470, 109086, -412749, 109498, -411841, 109862
        .AddBezierSegment -411841, 109862, -410933, 110226, -409938, 110496, -408858, 110673
        .AddBezierSegment -408858, 110673, -407418, 110910, -405757, 111029, -403876, 111029
        .AddBezierSegment -403876, 111029, -400468, 111029, -397749, 110617, -395701, 109795
        .AddBezierSegment -395701, 109795, -393659, 108971, -392637, 107826, -392637, 106360
        .AddBezierSegment -392637, 106360, -392637, 104906, -393653, 103768, -395680, 102948
        .AddBezierSegment -395680, 102948, -397706, 102126, -400420, 101716, -403811, 101716
        .AddBezierSegment -403811, 101716, -407246, 101716, -409982, 102124, -412008, 102942
        .AddBezierSegment -412008, 102942, -414034, 103757, -415050, 104883, -415050, 106318
        .CloseSubPath

        .StartSubPath
        .AddLineSegment -410245, 106379
        .AddBezierSegment -410245, 106379, -410245, 105360, -409638, 104588, -408428, 104061
        .AddBezierSegment -408428, 104061, -407219, 103536, -405682, 103272, -403822, 103272
        .AddBezierSegment -403822, 103272, -401957, 103272, -400431, 103534, -399232, 104054
        .AddBezierSegment -399232, 104054, -398039, 104577, -397442, 105360, -397442, 106404
        .AddBezierSegment -397442, 106404, -397442, 107435, -398023, 108205, -399189, 108711
        .AddBezierSegment -399189, 108711, -400350, 109219, -401893, 109473, -403822, 109473
        .AddBezierSegment -403822, 109473, -405746, 109473, -407300, 109215, -408477, 108703
        .AddBezierSegment -408477, 108703, -409654, 108188, -410245, 107414, -410245, 106379
        .CloseSubPath

        .StartSubPath
        .AddLineSegment -383807, 101879
        .AddLineSegment -383807, 109347
        .AddLineSegment -390644, 109347
        .AddLineSegment -390644, 110866
        .AddLineSegment -372327, 110866
        .AddLineSegment -372327, 109347
        .AddLineSegment -379131, 109347
        .AddLineSegment -379131, 101879
        .AddLineSegment -383807, 101879
        .CloseSubPath

        .StartSubPath
        .AddLineSegment -364109, 101879
        .AddLineSegment -364109, 109347
        .AddLineSegment -370945, 109347
        .AddLineSegment -370945, 110866
        .AddLineSegment -352628, 110866
        .AddLineSegment -352628, 109347
        .AddLineSegment -359433, 109347
        .AddLineSegment -359433, 101879
        .AddLineSegment -364109, 101879
        .CloseSubPath

        .StartSubPath
        .AddLineSegment -349603, 101879
        .AddLineSegment -349603, 110866
        .AddLineSegment -332511, 110866
        .AddLineSegment -332511, 109347
        .AddLineSegment -344926, 109347
        .AddLineSegment -344926, 107351
        .AddLineSegment -333382, 107351
        .AddLineSegment -333382, 105833
        .AddLineSegment -344926, 105833
        .AddLineSegment -344926, 103398
        .AddLineSegment -332060, 103398
        .AddLineSegment -332060, 101879
        .AddLineSegment -349603, 101879
        .CloseSubPath

        .StartSubPath
        .AddLineSegment -328093, 101879
        .AddLineSegment -328093, 110866
        .AddLineSegment -318273, 110866
        .AddBezierSegment -318273, 110866, -315807, 110866, -314011, 110784, -312894, 110623
        .AddBezierSegment -312894, 110623, -311770, 110462, -310878, 110175, -310206, 109763
        .AddBezierSegment -310206, 109763, -309534, 109349, -309196, 108879, -309196, 108347
        .AddBezierSegment -309196, 108347, -309196, 107671, -309701, 107115, -310717, 106676
        .AddBezierSegment -310717, 106676, -311733, 106236, -313248, 105960, -315269, 105845
        .AddBezierSegment -315269, 105845, -314264, 105617, -313436, 105366, -312786, 105092
        .AddBezierSegment -312786, 105092, -312130, 104818, -311249, 104331, -310142, 103632
        .AddLineSegment -307325, 101879
        .AddLineSegment -312899, 101879
        .AddLineSegment -316242, 103841
        .AddBezierSegment -316242, 103841, -317440, 104538, -318263, 104977, -318709, 105159
        .AddBezierSegment -318709, 105159, -319149, 105339, -319617, 105464, -320112, 105531
        .AddBezierSegment -320112, 105531, -320611, 105598, -321396, 105632, -322471, 105632
        .AddLineSegment -323417, 105632
        .AddLineSegment -323417, 101879
        .AddLineSegment -328093, 101879
        .CloseSubPath

        .StartSubPath
        .AddLineSegment -323417, 107063
        .AddLineSegment -319967, 107063
        .AddBezierSegment -319967, 107063, -317731, 107063, -316339, 107100, -315780, 107174
        .AddBezierSegment -315780, 107174, -315226, 107247, -314786, 107374, -314474, 107554
        .AddBezierSegment -314474, 107554, -314156, 107734, -314001, 107960, -314001, 108230
        .AddBezierSegment -314001, 108230, -314001, 108533, -314210, 108776, -314624, 108964
        .AddBezierSegment -314624, 108964, -315038, 109150, -315624, 109268, -316376, 109316
        .AddBezierSegment -316376, 109316, -316758, 109337, -317892, 109347, -319779, 109347
        .AddLineSegment -323417, 109347
        .AddLineSegment -323417, 107063
        .CloseSubPath

        .StartSubPath
        .AddLineSegment -281849, 108665
        .AddLineSegment -286139, 108481
        .AddBezierSegment -286139, 108481, -286240, 108824, -286515, 109077, -286955, 109240
        .AddBezierSegment -286955, 109240, -287396, 109403, -287966, 109485, -288670, 109485
        .AddBezierSegment -288670, 109485, -289605, 109485, -290395, 109322, -291040, 108996
        .AddBezierSegment -291040, 108996, -291685, 108669, -292088, 107989, -292255, 106956
        .AddBezierSegment -292255, 106956, -291153, 107462, -289788, 107715, -288149, 107715
        .AddBezierSegment -288149, 107715, -286305, 107715, -284719, 107441, -283403, 106893
        .AddBezierSegment -283403, 106893, -282086, 106343, -281430, 105636, -281430, 104768
        .AddBezierSegment -281430, 104768, -281430, 103847, -282123, 103107, -283510, 102550
        .AddBezierSegment -283510, 102550, -284897, 101994, -286676, 101716, -288847, 101716
        .AddBezierSegment -288847, 101716, -291174, 101716, -293093, 102069, -294593, 102776
        .AddBezierSegment -294593, 102776, -296092, 103483, -296845, 104640, -296845, 106251
        .AddBezierSegment -296845, 106251, -296845, 107901, -296060, 109092, -294496, 109822
        .AddBezierSegment -294496, 109822, -292932, 110552, -290900, 110916, -288401, 110916
        .AddBezierSegment -288401, 110916, -286649, 110916, -285198, 110724, -284048, 110339
        .AddBezierSegment -284048, 110339, -282897, 109956, -282166, 109397, -281849, 108665
        .CloseSubPath

        .StartSubPath
        .AddLineSegment -291879, 104904
        .AddBezierSegment -291879, 104904, -291879, 104343, -291545, 103910, -290884, 103605
        .AddBezierSegment -290884, 103605, -290223, 103299, -289471, 103147, -288616, 103147
        .AddBezierSegment -288616, 103147, -287799, 103147, -287117, 103272, -286568, 103521
        .AddBezierSegment -286568, 103521, -286025, 103772, -285751, 104182, -285751, 104751
        .AddBezierSegment -285751, 104751, -285751, 105337, -286047, 105766, -286633, 106040
        .AddBezierSegment -286633, 106040, -287224, 106312, -287955, 106448, -288836, 106448
        .AddBezierSegment -288836, 106448, -289691, 106448, -290406, 106318, -290997, 106056
        .AddBezierSegment -290997, 106056, -291583, 105797, -291879, 105412, -291879, 104904
        .CloseSubPath

        .StartSubPath
        .AddLineSegment -271412, 110916
        .AddBezierSegment -271412, 110916, -269176, 110916, -267429, 110604, -266166, 109981
        .AddBezierSegment -266166, 109981, -264667, 109245, -263914, 108021, -263914, 106314
        .AddBezierSegment -263914, 106314, -263914, 104609, -264672, 103383, -266182, 102638
        .AddBezierSegment -266182, 102638, -267435, 102023, -269176, 101716, -271412, 101716
        .AddBezierSegment -271412, 101716, -273659, 101716, -275470, 102052, -276846, 102726
        .AddBezierSegment -276846, 102726, -278222, 103400, -278909, 104603, -278909, 106331
        .AddBezierSegment -278909, 106331, -278909, 108027, -278152, 109249, -276641, 109993
        .AddBezierSegment -276641, 109993, -275389, 110609, -273648, 110916, -271412, 110916
        .CloseSubPath

        .StartSubPath
        .AddLineSegment -271412, 109485
        .AddBezierSegment -271412, 109485, -271939, 109485, -272412, 109418, -272820, 109284
        .AddBezierSegment -272820, 109284, -273234, 109153, -273556, 108914, -273782, 108569
        .AddBezierSegment -273782, 108569, -274083, 108123, -274233, 107370, -274233, 106314
        .AddBezierSegment -274233, 106314, -274233, 105255, -274099, 104529, -273830, 104134
        .AddBezierSegment -273830, 104134, -273562, 103739, -273223, 103475, -272815, 103343
        .AddBezierSegment -272815, 103343, -272406, 103211, -271939, 103147, -271412, 103147
        .AddBezierSegment -271412, 103147, -270885, 103147, -270412, 103213, -269998, 103347
        .AddBezierSegment -269998, 103347, -269590, 103479, -269267, 103718, -269042, 104063
        .AddBezierSegment -269042, 104063, -268741, 104504, -268590, 105255, -268590, 106314
        .AddBezierSegment -268590, 106314, -268590, 107370, -268725, 108096, -268993, 108492
        .AddBezierSegment -268993, 108492, -269262, 108887, -269601, 109153, -270009, 109284
        .AddBezierSegment -270009, 109284, -270417, 109418, -270885, 109485, -271412, 109485
        .CloseSubPath

        .StartSubPath
        .AddLineSegment -253477, 110916
        .AddBezierSegment -253477, 110916, -251241, 110916, -249494, 110604, -248231, 109981
        .AddBezierSegment -248231, 109981, -246732, 109245, -245979, 108021, -245979, 106314
        .AddBezierSegment -245979, 106314, -245979, 104609, -246737, 103383, -248247, 102638
        .AddBezierSegment -248247, 102638, -249499, 102023, -251241, 101716, -253477, 101716
        .AddBezierSegment -253477, 101716, -255723, 101716, -257535, 102052, -258910, 102726
        .AddBezierSegment -258910, 102726, -260286, 103400, -260974, 104603, -260974, 106331
        .AddBezierSegment -260974, 106331, -260974, 108027, -260217, 109249, -258706, 109993
        .AddBezierSegment -258706, 109993, -257454, 110609, -255713, 110916, -253477, 110916
        .CloseSubPath

        .StartSubPath
        .AddLineSegment -253477, 109485
        .AddBezierSegment -253477, 109485, -254004, 109485, -254476, 109418, -254885, 109284
        .AddBezierSegment -254885, 109284, -255299, 109153, -255621, 108914, -255847, 108569
        .AddBezierSegment -255847, 108569, -256148, 108123, -256298, 107370, -256298, 106314
        .AddBezierSegment -256298, 106314, -256298, 105255, -256164, 104529, -255895, 104134
        .AddBezierSegment -255895, 104134, -255626, 103739, -255288, 103475, -254880, 103343
        .AddBezierSegment -254880, 103343, -254471, 103211, -254004, 103147, -253477, 103147
        .AddBezierSegment -253477, 103147, -252950, 103147, -252477, 103213, -252063, 103347
        .AddBezierSegment -252063, 103347, -251655, 103479, -251332, 103718, -251106, 104063
        .AddBezierSegment -251106, 104063, -250806, 104504, -250655, 105255, -250655, 106314
        .AddBezierSegment -250655, 106314, -250655, 107370, -250789, 108096, -251058, 108492
        .AddBezierSegment -251058, 108492, -251327, 108887, -251666, 109153, -252074, 109284
        .AddBezierSegment -252074, 109284, -252482, 109418, -252950, 109485, -253477, 109485
        .CloseSubPath

        Set .brush = New waBrush
        .brush.SetLinearGradient -445000, 119000, -445000, 104000, RGB(43, 42, 41), RGB(89, 89, 89), WrapModeTileFlipXY
    End With
    MyCurves.Add path

    Set path = New waPath
    With path
        .StartSubPath
        .AddLineSegment -322315, 81964
        .AddLineSegment -322315, 88026
        .AddLineSegment -320352, 88026
        .AddBezierSegment -320352, 88026, -319608, 88026, -319124, 87997, -318899, 87936
        .AddBezierSegment -318899, 87936, -318551, 87844, -318261, 87647, -318026, 87342
        .AddBezierSegment -318026, 87342, -317792, 87037, -317675, 86643, -317675, 86161
        .AddBezierSegment -317675, 86161, -317675, 85788, -317743, 85476, -317877, 85222
        .AddBezierSegment -317877, 85222, -318012, 84968, -318183, 84769, -318390, 84624
        .AddBezierSegment -318390, 84624, -318598, 84480, -318810, 84384, -319024, 84338
        .AddBezierSegment -319024, 84338, -319316, 84280, -319738, 84250, -320291, 84250
        .AddLineSegment -321087, 84250
        .AddLineSegment -321087, 81964
        .AddLineSegment -322315, 81964
        .CloseSubPath

        .StartSubPath
        .AddLineSegment -321087, 87002
        .AddLineSegment -321087, 85275
        .AddLineSegment -320420, 85275
        .AddBezierSegment -320420, 85275, -319938, 85275, -319618, 85307, -319456, 85371
        .AddBezierSegment -319456, 85371, -319294, 85434, -319167, 85533, -319075, 85670
        .AddBezierSegment -319075, 85670, -318983, 85805, -318937, 85962, -318937, 86142
        .AddBezierSegment -318937, 86142, -318937, 86364, -319002, 86546, -319130, 86690
        .AddBezierSegment -319130, 86690, -319260, 86834, -319422, 86924, -319621, 86961
        .AddBezierSegment -319621, 86961, -319766, 86988, -320058, 87002, -320499, 87002
        .AddLineSegment -321087, 87002
        .CloseSubPath

        .StartSubPath
        .AddLineSegment -316634, 81964
        .AddLineSegment -316634, 87975
        .AddLineSegment -315406, 87975
        .AddLineSegment -315406, 82989
        .AddLineSegment -312366, 82989
        .AddLineSegment -312366, 81964
        .AddLineSegment -316634, 81964
        .CloseSubPath

        .StartSubPath
        .AddLineSegment -311750, 84959
        .AddBezierSegment -311750, 84959, -311750, 85577, -311657, 86096, -311472, 86516
        .AddBezierSegment -311472, 86516, -311335, 86825, -311146, 87103, -310907, 87349
        .AddBezierSegment -310907, 87349, -310669, 87594, -310408, 87776, -310124, 87896
        .AddBezierSegment -310124, 87896, -309746, 88056, -309310, 88136, -308816, 88136
        .AddBezierSegment -308816, 88136, -307921, 88136, -307207, 87858, -306670, 87304
        .AddBezierSegment -306670, 87304, -306134, 86748, -305865, 85976, -305865, 84987
        .AddBezierSegment -305865, 84987, -305865, 84006, -306132, 83238, -306664, 82685
        .AddBezierSegment -306664, 82685, -307196, 82131, -307909, 81854, -308799, 81854
        .AddBezierSegment -308799, 81854, -309701, 81854, -310419, 82129, -310951, 82681
        .AddBezierSegment -310951, 82681, -311483, 83231, -311750, 83990, -311750, 84959
        .CloseSubPath

        .StartSubPath
        .AddLineSegment -310488, 84999
        .AddBezierSegment -310488, 84999, -310488, 84312, -310329, 83791, -310011, 83436
        .AddBezierSegment -310011, 83436, -309694, 83082, -309290, 82904, -308802, 82904
        .AddBezierSegment -308802, 82904, -308312, 82904, -307912, 83080, -307597, 83432
        .AddBezierSegment -307597, 83432, -307284, 83784, -307127, 84312, -307127, 85016
        .AddBezierSegment -307127, 85016, -307127, 85712, -307279, 86231, -307586, 86573
        .AddBezierSegment -307586, 86573, -307890, 86916, -308295, 87087, -308802, 87087
        .AddBezierSegment -308802, 87087, -309307, 87087, -309715, 86913, -310024, 86567
        .AddBezierSegment -310024, 86567, -310333, 86220, -310488, 85698, -310488, 84999
        .CloseSubPath

        .StartSubPath
        .AddLineSegment -303547, 81964
        .AddLineSegment -303547, 87002
        .AddLineSegment -305342, 87002
        .AddLineSegment -305342, 88026
        .AddLineSegment -300533, 88026
        .AddLineSegment -300533, 87002
        .AddLineSegment -302319, 87002
        .AddLineSegment -302319, 81964
        .AddLineSegment -303547, 81964
        .CloseSubPath

        .StartSubPath
        .AddLineSegment -295861, 84199
        .AddLineSegment -294675, 83835
        .AddBezierSegment -294675, 83835, -294857, 83168, -295161, 82672, -295584, 82348
        .AddBezierSegment -295584, 82348, -296007, 82025, -296544, 81862, -297194, 81862
        .AddBezierSegment -297194, 81862, -298000, 81862, -298662, 82138, -299181, 82688
        .AddBezierSegment -299181, 82688, -299699, 83238, -299959, 83990, -299959, 84946
        .AddBezierSegment -299959, 84946, -299959, 85955, -299697, 86739, -299177, 87298
        .AddBezierSegment -299177, 87298, -298656, 87857, -297970, 88136, -297121, 88136
        .AddBezierSegment -297121, 88136, -296379, 88136, -295776, 87916, -295313, 87477
        .AddBezierSegment -295313, 87477, -295037, 87218, -294831, 86845, -294692, 86358
        .AddLineSegment -295903, 86070
        .AddBezierSegment -295903, 86070, -295975, 86384, -296125, 86632, -296352, 86814
        .AddBezierSegment -296352, 86814, -296580, 86996, -296857, 87087, -297182, 87087
        .AddBezierSegment -297182, 87087, -297632, 87087, -297997, 86926, -298276, 86602
        .AddBezierSegment -298276, 86602, -298557, 86281, -298697, 85759, -298697, 85036
        .AddBezierSegment -298697, 85036, -298697, 84271, -298559, 83725, -298284, 83401
        .AddBezierSegment -298284, 83401, -298007, 83075, -297649, 82912, -297207, 82912
        .AddBezierSegment -297207, 82912, -296881, 82912, -296600, 83015, -296366, 83223
        .AddBezierSegment -296366, 83223, -296132, 83430, -295962, 83755, -295861, 84199
        .CloseSubPath

        .StartSubPath
        .AddLineSegment -288163, 81964
        .AddLineSegment -289485, 81964
        .AddLineSegment -290009, 83344
        .AddLineSegment -292440, 83344
        .AddLineSegment -292944, 81964
        .AddLineSegment -294242, 81964
        .AddLineSegment -291888, 88026
        .AddLineSegment -290591, 88026
        .AddLineSegment -288163, 81964
        .CloseSubPath

        .StartSubPath
        .AddLineSegment -290407, 84369
        .AddLineSegment -291249, 86612
        .AddLineSegment -292068, 84369
        .AddLineSegment -290407, 84369
        .CloseSubPath

        .StartSubPath
        .AddLineSegment -287476, 81964
        .AddLineSegment -287476, 87975
        .AddLineSegment -286248, 87975
        .AddLineSegment -286248, 82989
        .AddLineSegment -283209, 82989
        .AddLineSegment -283209, 81964
        .AddLineSegment -287476, 81964
        .CloseSubPath

        .StartSubPath
        .AddLineSegment -278460, 84199
        .AddLineSegment -277275, 83835
        .AddBezierSegment -277275, 83835, -277457, 83168, -277760, 82672, -278184, 82348
        .AddBezierSegment -278184, 82348, -278607, 82025, -279143, 81862, -279794, 81862
        .AddBezierSegment -279794, 81862, -280599, 81862, -281261, 82138, -281781, 82688
        .AddBezierSegment -281781, 82688, -282298, 83238, -282558, 83990, -282558, 84946
        .AddBezierSegment -282558, 84946, -282558, 85955, -282297, 86739, -281776, 87298
        .AddBezierSegment -281776, 87298, -281256, 87857, -280570, 88136, -279720, 88136
        .AddBezierSegment -279720, 88136, -278978, 88136, -278375, 87916, -277913, 87477
        .AddBezierSegment -277913, 87477, -277636, 87218, -277430, 86845, -277292, 86358
        .AddLineSegment -278502, 86070
        .AddBezierSegment -278502, 86070, -278575, 86384, -278724, 86632, -278951, 86814
        .AddBezierSegment -278951, 86814, -279180, 86996, -279456, 87087, -279781, 87087
        .AddBezierSegment -279781, 87087, -280231, 87087, -280597, 86926, -280876, 86602
        .AddBezierSegment -280876, 86602, -281157, 86281, -281296, 85759, -281296, 85036
        .AddBezierSegment -281296, 85036, -281296, 84271, -281158, 83725, -280883, 83401
        .AddBezierSegment -280883, 83401, -280606, 83075, -280248, 82912, -279806, 82912
        .AddBezierSegment -279806, 82912, -279480, 82912, -279200, 83015, -278965, 83223
        .AddBezierSegment -278965, 83223, -278731, 83430, -278562, 83755, -278460, 84199
        .CloseSubPath

        Set .brush = New waBrush
        .brush.SetColor RGB(43, 42, 41), 100
    End With
    MyCurves.Add path
    
    Set path = New waPath
    With path
        .StartSubPath
        .AddLineSegment -13408, 81964
        .AddLineSegment -13408, 88026
        .AddLineSegment -11445, 88026
        .AddBezierSegment -11445, 88026, -10701, 88026, -10217, 87997, -9991, 87936
        .AddBezierSegment -9991, 87936, -9644, 87844, -9353, 87647, -9119, 87342
        .AddBezierSegment -9119, 87342, -8885, 87037, -8768, 86643, -8768, 86161
        .AddBezierSegment -8768, 86161, -8768, 85788, -8835, 85476, -8970, 85222
        .AddBezierSegment -8970, 85222, -9105, 84968, -9276, 84769, -9483, 84624
        .AddBezierSegment -9483, 84624, -9691, 84480, -9902, 84384, -10117, 84338
        .AddBezierSegment -10117, 84338, -10409, 84280, -10831, 84250, -11384, 84250
        .AddLineSegment -12180, 84250
        .AddLineSegment -12180, 81964
        .AddLineSegment -13408, 81964
        .CloseSubPath

        .StartSubPath
        .AddLineSegment -12180, 87002
        .AddLineSegment -12180, 85275
        .AddLineSegment -11512, 85275
        .AddBezierSegment -11512, 85275, -11031, 85275, -10711, 85307, -10549, 85371
        .AddBezierSegment -10549, 85371, -10386, 85434, -10259, 85533, -10168, 85670
        .AddBezierSegment -10168, 85670, -10076, 85805, -10029, 85962, -10029, 86142
        .AddBezierSegment -10029, 86142, -10029, 86364, -10094, 86546, -10223, 86690
        .AddBezierSegment -10223, 86690, -10352, 86834, -10515, 86924, -10714, 86961
        .AddBezierSegment -10714, 86961, -10859, 86988, -11151, 87002, -11591, 87002
        .AddLineSegment -12180, 87002
        .CloseSubPath

        .StartSubPath
        .AddLineSegment -7726, 81964
        .AddLineSegment -7726, 87975
        .AddLineSegment -6499, 87975
        .AddLineSegment -6499, 82989
        .AddLineSegment -3459, 82989
        .AddLineSegment -3459, 81964
        .AddLineSegment -7726, 81964
        .CloseSubPath

        .StartSubPath
        .AddLineSegment -2842, 84959
        .AddBezierSegment -2842, 84959, -2842, 85577, -2749, 86096, -2564, 86516
        .AddBezierSegment -2564, 86516, -2428, 86825, -2239, 87103, -2000, 87349
        .AddBezierSegment -2000, 87349, -1762, 87594, -1501, 87776, -1217, 87896
        .AddBezierSegment -1217, 87896, -839, 88056, -403, 88136, 91, 88136
        .AddBezierSegment 91, 88136, 986, 88136, 1700, 87858, 2237, 87304
        .AddBezierSegment 2237, 87304, 2774, 86748, 3042, 85976, 3042, 84987
        .AddBezierSegment 3042, 84987, 3042, 84006, 2775, 83238, 2243, 82685
        .AddBezierSegment 2243, 82685, 1711, 82131, 999, 81854, 108, 81854
        .AddBezierSegment 108, 81854, -794, 81854, -1512, 82129, -2044, 82681
        .AddBezierSegment -2044, 82681, -2576, 83231, -2842, 83990, -2842, 84959
        .CloseSubPath

        .StartSubPath
        .AddLineSegment -1581, 84999
        .AddBezierSegment -1581, 84999, -1581, 84312, -1421, 83791, -1104, 83436
        .AddBezierSegment -1104, 83436, -786, 83082, -383, 82904, 105, 82904
        .AddBezierSegment 105, 82904, 595, 82904, 996, 83080, 1310, 83432
        .AddBezierSegment 1310, 83432, 1624, 83784, 1780, 84312, 1780, 85016
        .AddBezierSegment 1780, 85016, 1780, 85712, 1628, 86231, 1322, 86573
        .AddBezierSegment 1322, 86573, 1017, 86916, 612, 87087, 105, 87087
        .AddBezierSegment 105, 87087, -400, 87087, -808, 86913, -1117, 86567
        .AddBezierSegment -1117, 86567, -1426, 86220, -1581, 85698, -1581, 84999
        .CloseSubPath

        .StartSubPath
        .AddLineSegment 5360, 81964
        .AddLineSegment 5360, 87002
        .AddLineSegment 3565, 87002
        .AddLineSegment 3565, 88026
        .AddLineSegment 8374, 88026
        .AddLineSegment 8374, 87002
        .AddLineSegment 6588, 87002
        .AddLineSegment 6588, 81964
        .AddLineSegment 5360, 81964
        .CloseSubPath

        .StartSubPath
        .AddLineSegment 13047, 84199
        .AddLineSegment 14232, 83835
        .AddBezierSegment 14232, 83835, 14050, 83168, 13746, 82672, 13323, 82348
        .AddBezierSegment 13323, 82348, 12900, 82025, 12364, 81862, 11713, 81862
        .AddBezierSegment 11713, 81862, 10907, 81862, 10246, 82138, 9726, 82688
        .AddBezierSegment 9726, 82688, 9208, 83238, 8949, 83990, 8949, 84946
        .AddBezierSegment 8949, 84946, 8949, 85955, 9210, 86739, 9730, 87298
        .AddBezierSegment 9730, 87298, 10251, 87857, 10937, 88136, 11787, 88136
        .AddBezierSegment 11787, 88136, 12529, 88136, 13131, 87916, 13594, 87477
        .AddBezierSegment 13594, 87477, 13871, 87218, 14077, 86845, 14215, 86358
        .AddLineSegment 13004, 86070
        .AddBezierSegment 13004, 86070, 12932, 86384, 12783, 86632, 12556, 86814
        .AddBezierSegment 12556, 86814, 12327, 86996, 12050, 87087, 11726, 87087
        .AddBezierSegment 11726, 87087, 11276, 87087, 10910, 86926, 10631, 86602
        .AddBezierSegment 10631, 86602, 10350, 86281, 10210, 85759, 10210, 85036
        .AddBezierSegment 10210, 85036, 10210, 84271, 10349, 83725, 10624, 83401
        .AddBezierSegment 10624, 83401, 10900, 83075, 11259, 82912, 11700, 82912
        .AddBezierSegment 11700, 82912, 12026, 82912, 12307, 83015, 12541, 83223
        .AddBezierSegment 12541, 83223, 12776, 83430, 12945, 83755, 13047, 84199
        .CloseSubPath

        .StartSubPath
        .AddLineSegment 20744, 81964
        .AddLineSegment 19422, 81964
        .AddLineSegment 18899, 83344
        .AddLineSegment 16467, 83344
        .AddLineSegment 15963, 81964
        .AddLineSegment 14665, 81964
        .AddLineSegment 17019, 88026
        .AddLineSegment 18316, 88026
        .AddLineSegment 20744, 81964
        .CloseSubPath

        .StartSubPath
        .AddLineSegment 18501, 84369
        .AddLineSegment 17658, 86612
        .AddLineSegment 16840, 84369
        .AddLineSegment 18501, 84369
        .CloseSubPath

        .StartSubPath
        .AddLineSegment 21431, 81964
        .AddLineSegment 21431, 87975
        .AddLineSegment 22659, 87975
        .AddLineSegment 22659, 82989
        .AddLineSegment 25699, 82989
        .AddLineSegment 25699, 81964
        .AddLineSegment 21431, 81964
        .CloseSubPath

        .StartSubPath
        .AddLineSegment 30447, 84199
        .AddLineSegment 31633, 83835
        .AddBezierSegment 31633, 83835, 31450, 83168, 31147, 82672, 30724, 82348
        .AddBezierSegment 30724, 82348, 30300, 82025, 29764, 81862, 29114, 81862
        .AddBezierSegment 29114, 81862, 28308, 81862, 27646, 82138, 27127, 82688
        .AddBezierSegment 27127, 82688, 26609, 83238, 26349, 83990, 26349, 84946
        .AddBezierSegment 26349, 84946, 26349, 85955, 26610, 86739, 27131, 87298
        .AddBezierSegment 27131, 87298, 27652, 87857, 28337, 88136, 29187, 88136
        .AddBezierSegment 29187, 88136, 29929, 88136, 30532, 87916, 30995, 87477
        .AddBezierSegment 30995, 87477, 31271, 87218, 31477, 86845, 31616, 86358
        .AddLineSegment 30405, 86070
        .AddBezierSegment 30405, 86070, 30333, 86384, 30183, 86632, 29956, 86814
        .AddBezierSegment 29956, 86814, 29727, 86996, 29451, 87087, 29126, 87087
        .AddBezierSegment 29126, 87087, 28676, 87087, 28311, 86926, 28031, 86602
        .AddBezierSegment 28031, 86602, 27750, 86281, 27611, 85759, 27611, 85036
        .AddBezierSegment 27611, 85036, 27611, 84271, 27749, 83725, 28024, 83401
        .AddBezierSegment 28024, 83401, 28301, 83075, 28659, 82912, 29101, 82912
        .AddBezierSegment 29101, 82912, 29427, 82912, 29708, 83015, 29942, 83223
        .AddBezierSegment 29942, 83223, 30176, 83430, 30346, 83755, 30447, 84199
        .CloseSubPath

        Set .brush = New waBrush
        .brush.SetColor RGB(43, 42, 41), 100
    End With
    MyCurves.Add path

    Set path = New waPath
    With path
        .StartSubPath
        .AddLineSegment -634982, 81964
        .AddLineSegment -634982, 88026
        .AddLineSegment -633019, 88026
        .AddBezierSegment -633019, 88026, -632275, 88026, -631791, 87997, -631566, 87936
        .AddBezierSegment -631566, 87936, -631218, 87844, -630928, 87647, -630693, 87342
        .AddBezierSegment -630693, 87342, -630459, 87037, -630342, 86643, -630342, 86161
        .AddBezierSegment -630342, 86161, -630342, 85788, -630410, 85476, -630544, 85222
        .AddBezierSegment -630544, 85222, -630679, 84968, -630850, 84769, -631058, 84624
        .AddBezierSegment -631058, 84624, -631265, 84480, -631477, 84384, -631691, 84338
        .AddBezierSegment -631691, 84338, -631983, 84280, -632405, 84250, -632958, 84250
        .AddLineSegment -633754, 84250
        .AddLineSegment -633754, 81964
        .AddLineSegment -634982, 81964
        .CloseSubPath

        .StartSubPath
        .AddLineSegment -633754, 87002
        .AddLineSegment -633754, 85275
        .AddLineSegment -633087, 85275
        .AddBezierSegment -633087, 85275, -632606, 85275, -632285, 85307, -632123, 85371
        .AddBezierSegment -632123, 85371, -631961, 85434, -631834, 85533, -631742, 85670
        .AddBezierSegment -631742, 85670, -631650, 85805, -631604, 85962, -631604, 86142
        .AddBezierSegment -631604, 86142, -631604, 86364, -631668, 86546, -631797, 86690
        .AddBezierSegment -631797, 86690, -631927, 86834, -632089, 86924, -632288, 86961
        .AddBezierSegment -632288, 86961, -632433, 86988, -632726, 87002, -633166, 87002
        .AddLineSegment -633754, 87002
        .CloseSubPath

        .StartSubPath
        .AddLineSegment -629301, 81964
        .AddLineSegment -629301, 87975
        .AddLineSegment -628073, 87975
        .AddLineSegment -628073, 82989
        .AddLineSegment -625033, 82989
        .AddLineSegment -625033, 81964
        .AddLineSegment -629301, 81964
        .CloseSubPath

        .StartSubPath
        .AddLineSegment -624417, 84959
        .AddBezierSegment -624417, 84959, -624417, 85577, -624324, 86096, -624139, 86516
        .AddBezierSegment -624139, 86516, -624002, 86825, -623813, 87103, -623574, 87349
        .AddBezierSegment -623574, 87349, -623336, 87594, -623075, 87776, -622791, 87896
        .AddBezierSegment -622791, 87896, -622413, 88056, -621977, 88136, -621483, 88136
        .AddBezierSegment -621483, 88136, -620588, 88136, -619874, 87858, -619337, 87304
        .AddBezierSegment -619337, 87304, -618801, 86748, -618533, 85976, -618533, 84987
        .AddBezierSegment -618533, 84987, -618533, 84006, -618799, 83238, -619331, 82685
        .AddBezierSegment -619331, 82685, -619863, 82131, -620576, 81854, -621466, 81854
        .AddBezierSegment -621466, 81854, -622368, 81854, -623086, 82129, -623618, 82681
        .AddBezierSegment -623618, 82681, -624150, 83231, -624417, 83990, -624417, 84959
        .CloseSubPath

        .StartSubPath
        .AddLineSegment -623155, 84999
        .AddBezierSegment -623155, 84999, -623155, 84312, -622996, 83791, -622678, 83436
        .AddBezierSegment -622678, 83436, -622361, 83082, -621957, 82904, -621469, 82904
        .AddBezierSegment -621469, 82904, -620979, 82904, -620579, 83080, -620264, 83432
        .AddBezierSegment -620264, 83432, -619951, 83784, -619794, 84312, -619794, 85016
        .AddBezierSegment -619794, 85016, -619794, 85712, -619946, 86231, -620253, 86573
        .AddBezierSegment -620253, 86573, -620557, 86916, -620962, 87087, -621469, 87087
        .AddBezierSegment -621469, 87087, -621974, 87087, -622382, 86913, -622691, 86567
        .AddBezierSegment -622691, 86567, -623000, 86220, -623155, 85698, -623155, 84999
        .CloseSubPath

        .StartSubPath
        .AddLineSegment -616214, 81964
        .AddLineSegment -616214, 87002
        .AddLineSegment -618009, 87002
        .AddLineSegment -618009, 88026
        .AddLineSegment -613200, 88026
        .AddLineSegment -613200, 87002
        .AddLineSegment -614986, 87002
        .AddLineSegment -614986, 81964
        .AddLineSegment -616214, 81964
        .CloseSubPath

        .StartSubPath
        .AddLineSegment -608528, 84199
        .AddLineSegment -607342, 83835
        .AddBezierSegment -607342, 83835, -607524, 83168, -607828, 82672, -608251, 82348
        .AddBezierSegment -608251, 82348, -608674, 82025, -609211, 81862, -609861, 81862
        .AddBezierSegment -609861, 81862, -610667, 81862, -611329, 82138, -611848, 82688
        .AddBezierSegment -611848, 82688, -612366, 83238, -612626, 83990, -612626, 84946
        .AddBezierSegment -612626, 84946, -612626, 85955, -612365, 86739, -611844, 87298
        .AddBezierSegment -611844, 87298, -611323, 87857, -610637, 88136, -609788, 88136
        .AddBezierSegment -609788, 88136, -609046, 88136, -608443, 87916, -607980, 87477
        .AddBezierSegment -607980, 87477, -607704, 87218, -607498, 86845, -607359, 86358
        .AddLineSegment -608570, 86070
        .AddBezierSegment -608570, 86070, -608642, 86384, -608792, 86632, -609019, 86814
        .AddBezierSegment -609019, 86814, -609247, 86996, -609524, 87087, -609848, 87087
        .AddBezierSegment -609848, 87087, -610299, 87087, -610664, 86926, -610943, 86602
        .AddBezierSegment -610943, 86602, -611224, 86281, -611364, 85759, -611364, 85036
        .AddBezierSegment -611364, 85036, -611364, 84271, -611226, 83725, -610951, 83401
        .AddBezierSegment -610951, 83401, -610674, 83075, -610316, 82912, -609874, 82912
        .AddBezierSegment -609874, 82912, -609548, 82912, -609267, 83015, -609033, 83223
        .AddBezierSegment -609033, 83223, -608799, 83430, -608629, 83755, -608528, 84199
        .CloseSubPath

        .StartSubPath
        .AddLineSegment -600830, 81964
        .AddLineSegment -602152, 81964
        .AddLineSegment -602676, 83344
        .AddLineSegment -605107, 83344
        .AddLineSegment -605611, 81964
        .AddLineSegment -606909, 81964
        .AddLineSegment -604555, 88026
        .AddLineSegment -603259, 88026
        .AddLineSegment -600830, 81964
        .CloseSubPath

        .StartSubPath
        .AddLineSegment -603074, 84369
        .AddLineSegment -603916, 86612
        .AddLineSegment -604735, 84369
        .AddLineSegment -603074, 84369
        .CloseSubPath

        .StartSubPath
        .AddLineSegment -600143, 81964
        .AddLineSegment -600143, 87975
        .AddLineSegment -598915, 87975
        .AddLineSegment -598915, 82989
        .AddLineSegment -595875, 82989
        .AddLineSegment -595875, 81964
        .AddLineSegment -600143, 81964
        .CloseSubPath

        .StartSubPath
        .AddLineSegment -591127, 84199
        .AddLineSegment -589942, 83835
        .AddBezierSegment -589942, 83835, -590124, 83168, -590427, 82672, -590851, 82348
        .AddBezierSegment -590851, 82348, -591274, 82025, -591810, 81862, -592461, 81862
        .AddBezierSegment -592461, 81862, -593266, 81862, -593928, 82138, -594447, 82688
        .AddBezierSegment -594447, 82688, -594965, 83238, -595225, 83990, -595225, 84946
        .AddBezierSegment -595225, 84946, -595225, 85955, -594964, 86739, -594443, 87298
        .AddBezierSegment -594443, 87298, -593923, 87857, -593237, 88136, -592387, 88136
        .AddBezierSegment -592387, 88136, -591645, 88136, -591043, 87916, -590580, 87477
        .AddBezierSegment -590580, 87477, -590303, 87218, -590097, 86845, -589959, 86358
        .AddLineSegment -591169, 86070
        .AddBezierSegment -591169, 86070, -591241, 86384, -591391, 86632, -591618, 86814
        .AddBezierSegment -591618, 86814, -591847, 86996, -592123, 87087, -592448, 87087
        .AddBezierSegment -592448, 87087, -592898, 87087, -593264, 86926, -593543, 86602
        .AddBezierSegment -593543, 86602, -593824, 86281, -593963, 85759, -593963, 85036
        .AddBezierSegment -593963, 85036, -593963, 84271, -593825, 83725, -593550, 83401
        .AddBezierSegment -593550, 83401, -593274, 83075, -592915, 82912, -592473, 82912
        .AddBezierSegment -592473, 82912, -592147, 82912, -591867, 83015, -591632, 83223
        .AddBezierSegment -591632, 83223, -591398, 83430, -591229, 83755, -591127, 84199
        .CloseSubPath

        Set .brush = New waBrush
        .brush.SetColor RGB(43, 42, 41), 100
    End With
    MyCurves.Add path

    Set path = New waPath
    With path
        .StartSubPath
        .AddLineSegment -164383, 81964
        .AddLineSegment -164383, 88026
        .AddLineSegment -162420, 88026
        .AddBezierSegment -162420, 88026, -161677, 88026, -161193, 87997, -160967, 87936
        .AddBezierSegment -160967, 87936, -160620, 87844, -160329, 87647, -160095, 87342
        .AddBezierSegment -160095, 87342, -159861, 87037, -159743, 86643, -159743, 86161
        .AddBezierSegment -159743, 86161, -159743, 85788, -159811, 85476, -159945, 85222
        .AddBezierSegment -159945, 85222, -160081, 84968, -160251, 84769, -160459, 84624
        .AddBezierSegment -160459, 84624, -160666, 84480, -160878, 84384, -161092, 84338
        .AddBezierSegment -161092, 84338, -161385, 84280, -161806, 84250, -162360, 84250
        .AddLineSegment -163155, 84250
        .AddLineSegment -163155, 81964
        .AddLineSegment -164383, 81964
        .CloseSubPath

        .StartSubPath
        .AddLineSegment -163155, 87002
        .AddLineSegment -163155, 85275
        .AddLineSegment -162488, 85275
        .AddBezierSegment -162488, 85275, -162007, 85275, -161686, 85307, -161524, 85371
        .AddBezierSegment -161524, 85371, -161362, 85434, -161235, 85533, -161143, 85670
        .AddBezierSegment -161143, 85670, -161052, 85805, -161005, 85962, -161005, 86142
        .AddBezierSegment -161005, 86142, -161005, 86364, -161070, 86546, -161198, 86690
        .AddBezierSegment -161198, 86690, -161328, 86834, -161490, 86924, -161689, 86961
        .AddBezierSegment -161689, 86961, -161835, 86988, -162127, 87002, -162567, 87002
        .AddLineSegment -163155, 87002
        .CloseSubPath

        .StartSubPath
        .AddLineSegment -158702, 81964
        .AddLineSegment -158702, 87975
        .AddLineSegment -157474, 87975
        .AddLineSegment -157474, 82989
        .AddLineSegment -154435, 82989
        .AddLineSegment -154435, 81964
        .AddLineSegment -158702, 81964
        .CloseSubPath

        .StartSubPath
        .AddLineSegment -153818, 84959
        .AddBezierSegment -153818, 84959, -153818, 85577, -153725, 86096, -153540, 86516
        .AddBezierSegment -153540, 86516, -153403, 86825, -153214, 87103, -152976, 87349
        .AddBezierSegment -152976, 87349, -152737, 87594, -152476, 87776, -152192, 87896
        .AddBezierSegment -152192, 87896, -151814, 88056, -151378, 88136, -150884, 88136
        .AddBezierSegment -150884, 88136, -149990, 88136, -149276, 87858, -148738, 87304
        .AddBezierSegment -148738, 87304, -148202, 86748, -147934, 85976, -147934, 84987
        .AddBezierSegment -147934, 84987, -147934, 84006, -148201, 83238, -148732, 82685
        .AddBezierSegment -148732, 82685, -149264, 82131, -149977, 81854, -150868, 81854
        .AddBezierSegment -150868, 81854, -151769, 81854, -152487, 82129, -153019, 82681
        .AddBezierSegment -153019, 82681, -153551, 83231, -153818, 83990, -153818, 84959
        .CloseSubPath

        .StartSubPath
        .AddLineSegment -152557, 84999
        .AddBezierSegment -152557, 84999, -152557, 84312, -152397, 83791, -152080, 83436
        .AddBezierSegment -152080, 83436, -151762, 83082, -151359, 82904, -150870, 82904
        .AddBezierSegment -150870, 82904, -150381, 82904, -149980, 83080, -149665, 83432
        .AddBezierSegment -149665, 83432, -149352, 83784, -149195, 84312, -149195, 85016
        .AddBezierSegment -149195, 85016, -149195, 85712, -149348, 86231, -149654, 86573
        .AddBezierSegment -149654, 86573, -149959, 86916, -150364, 87087, -150870, 87087
        .AddBezierSegment -150870, 87087, -151375, 87087, -151783, 86913, -152092, 86567
        .AddBezierSegment -152092, 86567, -152401, 86220, -152557, 85698, -152557, 84999
        .CloseSubPath

        .StartSubPath
        .AddLineSegment -145615, 81964
        .AddLineSegment -145615, 87002
        .AddLineSegment -147410, 87002
        .AddLineSegment -147410, 88026
        .AddLineSegment -142601, 88026
        .AddLineSegment -142601, 87002
        .AddLineSegment -144388, 87002
        .AddLineSegment -144388, 81964
        .AddLineSegment -145615, 81964
        .CloseSubPath

        .StartSubPath
        .AddLineSegment -137929, 84199
        .AddLineSegment -136744, 83835
        .AddBezierSegment -136744, 83835, -136926, 83168, -137229, 82672, -137652, 82348
        .AddBezierSegment -137652, 82348, -138076, 82025, -138612, 81862, -139262, 81862
        .AddBezierSegment -139262, 81862, -140068, 81862, -140730, 82138, -141249, 82688
        .AddBezierSegment -141249, 82688, -141767, 83238, -142027, 83990, -142027, 84946
        .AddBezierSegment -142027, 84946, -142027, 85955, -141766, 86739, -141245, 87298
        .AddBezierSegment -141245, 87298, -140724, 87857, -140039, 88136, -139189, 88136
        .AddBezierSegment -139189, 88136, -138447, 88136, -137844, 87916, -137381, 87477
        .AddBezierSegment -137381, 87477, -137105, 87218, -136899, 86845, -136761, 86358
        .AddLineSegment -137971, 86070
        .AddBezierSegment -137971, 86070, -138043, 86384, -138193, 86632, -138420, 86814
        .AddBezierSegment -138420, 86814, -138649, 86996, -138925, 87087, -139250, 87087
        .AddBezierSegment -139250, 87087, -139700, 87087, -140065, 86926, -140345, 86602
        .AddBezierSegment -140345, 86602, -140626, 86281, -140765, 85759, -140765, 85036
        .AddBezierSegment -140765, 85036, -140765, 84271, -140627, 83725, -140352, 83401
        .AddBezierSegment -140352, 83401, -140075, 83075, -139717, 82912, -139275, 82912
        .AddBezierSegment -139275, 82912, -138949, 82912, -138668, 83015, -138434, 83223
        .AddBezierSegment -138434, 83223, -138200, 83430, -138031, 83755, -137929, 84199
        .CloseSubPath

        .StartSubPath
        .AddLineSegment -130231, 81964
        .AddLineSegment -131554, 81964
        .AddLineSegment -132077, 83344
        .AddLineSegment -134508, 83344
        .AddLineSegment -135012, 81964
        .AddLineSegment -136310, 81964
        .AddLineSegment -133957, 88026
        .AddLineSegment -132660, 88026
        .AddLineSegment -130231, 81964
        .CloseSubPath

        .StartSubPath
        .AddLineSegment -132475, 84369
        .AddLineSegment -133317, 86612
        .AddLineSegment -134136, 84369
        .AddLineSegment -132475, 84369
        .CloseSubPath

        .StartSubPath
        .AddLineSegment -129544, 81964
        .AddLineSegment -129544, 87975
        .AddLineSegment -128316, 87975
        .AddLineSegment -128316, 82989
        .AddLineSegment -125277, 82989
        .AddLineSegment -125277, 81964
        .AddLineSegment -129544, 81964
        .CloseSubPath

        .StartSubPath
        .AddLineSegment -120528, 84199
        .AddLineSegment -119343, 83835
        .AddBezierSegment -119343, 83835, -119525, 83168, -119829, 82672, -120252, 82348
        .AddBezierSegment -120252, 82348, -120675, 82025, -121212, 81862, -121862, 81862
        .AddBezierSegment -121862, 81862, -122668, 81862, -123329, 82138, -123849, 82688
        .AddBezierSegment -123849, 82688, -124367, 83238, -124626, 83990, -124626, 84946
        .AddBezierSegment -124626, 84946, -124626, 85955, -124365, 86739, -123845, 87298
        .AddBezierSegment -123845, 87298, -123324, 87857, -122638, 88136, -121789, 88136
        .AddBezierSegment -121789, 88136, -121046, 88136, -120444, 87916, -119981, 87477
        .AddBezierSegment -119981, 87477, -119704, 87218, -119498, 86845, -119360, 86358
        .AddLineSegment -120571, 86070
        .AddBezierSegment -120571, 86070, -120643, 86384, -120792, 86632, -121020, 86814
        .AddBezierSegment -121020, 86814, -121248, 86996, -121525, 87087, -121849, 87087
        .AddBezierSegment -121849, 87087, -122299, 87087, -122665, 86926, -122944, 86602
        .AddBezierSegment -122944, 86602, -123225, 86281, -123365, 85759, -123365, 85036
        .AddBezierSegment -123365, 85036, -123365, 84271, -123227, 83725, -122951, 83401
        .AddBezierSegment -122951, 83401, -122675, 83075, -122316, 82912, -121875, 82912
        .AddBezierSegment -121875, 82912, -121549, 82912, -121268, 83015, -121034, 83223
        .AddBezierSegment -121034, 83223, -120799, 83430, -120630, 83755, -120528, 84199
        .CloseSubPath

        Set .brush = New waBrush
        .brush.SetColor RGB(43, 42, 41), 100
    End With
    MyCurves.Add path

    Set path = New waPath
    With path
        .StartSubPath
        .AddLineSegment -484442, 81964
        .AddLineSegment -484442, 88026
        .AddLineSegment -482479, 88026
        .AddBezierSegment -482479, 88026, -481735, 88026, -481251, 87997, -481025, 87936
        .AddBezierSegment -481025, 87936, -480678, 87844, -480387, 87647, -480153, 87342
        .AddBezierSegment -480153, 87342, -479919, 87037, -479802, 86643, -479802, 86161
        .AddBezierSegment -479802, 86161, -479802, 85788, -479870, 85476, -480004, 85222
        .AddBezierSegment -480004, 85222, -480139, 84968, -480310, 84769, -480517, 84624
        .AddBezierSegment -480517, 84624, -480725, 84480, -480936, 84384, -481151, 84338
        .AddBezierSegment -481151, 84338, -481443, 84280, -481865, 84250, -482418, 84250
        .AddLineSegment -483214, 84250
        .AddLineSegment -483214, 81964
        .AddLineSegment -484442, 81964
        .CloseSubPath

        .StartSubPath
        .AddLineSegment -483214, 87002
        .AddLineSegment -483214, 85275
        .AddLineSegment -482546, 85275
        .AddBezierSegment -482546, 85275, -482065, 85275, -481745, 85307, -481583, 85371
        .AddBezierSegment -481583, 85371, -481420, 85434, -481293, 85533, -481202, 85670
        .AddBezierSegment -481202, 85670, -481110, 85805, -481063, 85962, -481063, 86142
        .AddBezierSegment -481063, 86142, -481063, 86364, -481128, 86546, -481257, 86690
        .AddBezierSegment -481257, 86690, -481387, 86834, -481549, 86924, -481748, 86961
        .AddBezierSegment -481748, 86961, -481893, 86988, -482185, 87002, -482625, 87002
        .AddLineSegment -483214, 87002
        .CloseSubPath

        .StartSubPath
        .AddLineSegment -478760, 81964
        .AddLineSegment -478760, 87975
        .AddLineSegment -477533, 87975
        .AddLineSegment -477533, 82989
        .AddLineSegment -474493, 82989
        .AddLineSegment -474493, 81964
        .AddLineSegment -478760, 81964
        .CloseSubPath

        .StartSubPath
        .AddLineSegment -473877, 84959
        .AddBezierSegment -473877, 84959, -473877, 85577, -473783, 86096, -473599, 86516
        .AddBezierSegment -473599, 86516, -473462, 86825, -473273, 87103, -473034, 87349
        .AddBezierSegment -473034, 87349, -472796, 87594, -472535, 87776, -472251, 87896
        .AddBezierSegment -472251, 87896, -471873, 88056, -471437, 88136, -470943, 88136
        .AddBezierSegment -470943, 88136, -470048, 88136, -469334, 87858, -468796, 87304
        .AddBezierSegment -468796, 87304, -468260, 86748, -467992, 85976, -467992, 84987
        .AddBezierSegment -467992, 84987, -467992, 84006, -468259, 83238, -468791, 82685
        .AddBezierSegment -468791, 82685, -469323, 82131, -470036, 81854, -470926, 81854
        .AddBezierSegment -470926, 81854, -471828, 81854, -472546, 82129, -473078, 82681
        .AddBezierSegment -473078, 82681, -473610, 83231, -473877, 83990, -473877, 84959
        .CloseSubPath

        .StartSubPath
        .AddLineSegment -472615, 84999
        .AddBezierSegment -472615, 84999, -472615, 84312, -472455, 83791, -472138, 83436
        .AddBezierSegment -472138, 83436, -471820, 83082, -471417, 82904, -470929, 82904
        .AddBezierSegment -470929, 82904, -470439, 82904, -470038, 83080, -469724, 83432
        .AddBezierSegment -469724, 83432, -469410, 83784, -469254, 84312, -469254, 85016
        .AddBezierSegment -469254, 85016, -469254, 85712, -469406, 86231, -469712, 86573
        .AddBezierSegment -469712, 86573, -470017, 86916, -470422, 87087, -470929, 87087
        .AddBezierSegment -470929, 87087, -471434, 87087, -471842, 86913, -472151, 86567
        .AddBezierSegment -472151, 86567, -472460, 86220, -472615, 85698, -472615, 84999
        .CloseSubPath

        .StartSubPath
        .AddLineSegment -465674, 81964
        .AddLineSegment -465674, 87002
        .AddLineSegment -467469, 87002
        .AddLineSegment -467469, 88026
        .AddLineSegment -462659, 88026
        .AddLineSegment -462659, 87002
        .AddLineSegment -464446, 87002
        .AddLineSegment -464446, 81964
        .AddLineSegment -465674, 81964
        .CloseSubPath

        .StartSubPath
        .AddLineSegment -457987, 84199
        .AddLineSegment -456802, 83835
        .AddBezierSegment -456802, 83835, -456984, 83168, -457288, 82672, -457711, 82348
        .AddBezierSegment -457711, 82348, -458134, 82025, -458670, 81862, -459321, 81862
        .AddBezierSegment -459321, 81862, -460127, 81862, -460788, 82138, -461308, 82688
        .AddBezierSegment -461308, 82688, -461826, 83238, -462085, 83990, -462085, 84946
        .AddBezierSegment -462085, 84946, -462085, 85955, -461824, 86739, -461303, 87298
        .AddBezierSegment -461303, 87298, -460783, 87857, -460097, 88136, -459248, 88136
        .AddBezierSegment -459248, 88136, -458505, 88136, -457903, 87916, -457440, 87477
        .AddBezierSegment -457440, 87477, -457163, 87218, -456957, 86845, -456819, 86358
        .AddLineSegment -458030, 86070
        .AddBezierSegment -458030, 86070, -458102, 86384, -458251, 86632, -458478, 86814
        .AddBezierSegment -458478, 86814, -458707, 86996, -458984, 87087, -459308, 87087
        .AddBezierSegment -459308, 87087, -459758, 87087, -460124, 86926, -460403, 86602
        .AddBezierSegment -460403, 86602, -460684, 86281, -460824, 85759, -460824, 85036
        .AddBezierSegment -460824, 85036, -460824, 84271, -460685, 83725, -460410, 83401
        .AddBezierSegment -460410, 83401, -460134, 83075, -459775, 82912, -459334, 82912
        .AddBezierSegment -459334, 82912, -459008, 82912, -458727, 83015, -458493, 83223
        .AddBezierSegment -458493, 83223, -458258, 83430, -458089, 83755, -457987, 84199
        .CloseSubPath

        .StartSubPath
        .AddLineSegment -450290, 81964
        .AddLineSegment -451612, 81964
        .AddLineSegment -452135, 83344
        .AddLineSegment -454567, 83344
        .AddLineSegment -455071, 81964
        .AddLineSegment -456369, 81964
        .AddLineSegment -454015, 88026
        .AddLineSegment -452718, 88026
        .AddLineSegment -450290, 81964
        .CloseSubPath

        .StartSubPath
        .AddLineSegment -452533, 84369
        .AddLineSegment -453376, 86612
        .AddLineSegment -454194, 84369
        .AddLineSegment -452533, 84369
        .CloseSubPath

        .StartSubPath
        .AddLineSegment -449603, 81964
        .AddLineSegment -449603, 87975
        .AddLineSegment -448375, 87975
        .AddLineSegment -448375, 82989
        .AddLineSegment -445335, 82989
        .AddLineSegment -445335, 81964
        .AddLineSegment -449603, 81964
        .CloseSubPath

        .StartSubPath
        .AddLineSegment -440587, 84199
        .AddLineSegment -439401, 83835
        .AddBezierSegment -439401, 83835, -439584, 83168, -439887, 82672, -440310, 82348
        .AddBezierSegment -440310, 82348, -440734, 82025, -441270, 81862, -441920, 81862
        .AddBezierSegment -441920, 81862, -442726, 81862, -443388, 82138, -443907, 82688
        .AddBezierSegment -443907, 82688, -444425, 83238, -444685, 83990, -444685, 84946
        .AddBezierSegment -444685, 84946, -444685, 85955, -444424, 86739, -443903, 87298
        .AddBezierSegment -443903, 87298, -443382, 87857, -442697, 88136, -441847, 88136
        .AddBezierSegment -441847, 88136, -441105, 88136, -440502, 87916, -440039, 87477
        .AddBezierSegment -440039, 87477, -439763, 87218, -439557, 86845, -439418, 86358
        .AddLineSegment -440629, 86070
        .AddBezierSegment -440629, 86070, -440701, 86384, -440851, 86632, -441078, 86814
        .AddBezierSegment -441078, 86814, -441307, 86996, -441583, 87087, -441908, 87087
        .AddBezierSegment -441908, 87087, -442358, 87087, -442723, 86926, -443003, 86602
        .AddBezierSegment -443003, 86602, -443284, 86281, -443423, 85759, -443423, 85036
        .AddBezierSegment -443423, 85036, -443423, 84271, -443285, 83725, -443010, 83401
        .AddBezierSegment -443010, 83401, -442733, 83075, -442375, 82912, -441933, 82912
        .AddBezierSegment -441933, 82912, -441607, 82912, -441326, 83015, -441092, 83223
        .AddBezierSegment -441092, 83223, -440858, 83430, -440688, 83755, -440587, 84199
        .CloseSubPath

        Set .brush = New waBrush
        .brush.SetColor RGB(43, 42, 41), 100
    End With
    MyCurves.Add path

    Set path = New waPath
    With path
        .StartSubPath
        .AddLineSegment -670456, 119627
        .AddLineSegment -732318, 115076
        .AddBezierSegment -732318, 115076, -762591, 89710, -756994, -9203, -754739, -50659
        .AddLineSegment -670456, -57346
        .AddLineSegment -670456, 119627
        .CloseSubPath

        Set .brush = New waBrush
        .brush.SetLinearGradient -714000, 59000, -714000, 236000, RGB(228, 239, 240), RGB(177, 221, 227), WrapModeTileFlipXY
    End With
    MyCurves.Add path

    Set path = New waPath
    With path
        .StartSubPath
        .AddLineSegment 70541, 119627
        .AddLineSegment 132402, 115076
        .AddBezierSegment 132402, 115076, 162676, 89710, 157079, -9203, 154824, -50659
        .AddLineSegment 70541, -57346
        .AddLineSegment 70541, 119627
        .CloseSubPath

        Set .brush = New waBrush
        .brush.SetLinearGradient 114000, 59000, 114000, 236000, RGB(228, 239, 240), RGB(177, 221, 227), WrapModeTileFlipXY
    End With
    MyCurves.Add path
End Sub

' -----------------------------------------------------------------------------
' InitStandardPlotter1600_2DModel (Sub), source lines 3977-7110
' -----------------------------------------------------------------------------
Private Sub InitStandardPlotter1600_2DModel()
    Dim MyCurves As New Collection
    Set StandardPlotter1600_2DModel.Curves = MyCurves
    Dim path As waPath

    Set path = New waPath
    With path
        .StartSubPath
        .AddLineSegment 70137, -770784
        .AddLineSegment 69429, -756279
        .AddLineSegment 88532, -756279
        .AddLineSegment 87897, -770784
        .AddBezierSegment 87897, -770784, 84271, -771828, 73629, -772120, 70137, -770784
        .CloseSubPath

        Set .brush = New waBrush
        .brush.SetLinearGradient 60000, -764000, 79000, -764000, RGB(67, 66, 66), RGB(143, 143, 143), WrapModeTileFlipXY
    End With
    MyCurves.Add path

    Set path = New waPath
    With path
        .StartSubPath
        .AddLineSegment -14953, -856615
        .AddLineSegment -15661, -842109
        .AddLineSegment 3442, -842109
        .AddLineSegment 2807, -856615
        .AddBezierSegment 2807, -856615, -819, -857659, -11461, -857951, -14953, -856615
        .CloseSubPath

        Set .brush = New waBrush
        .brush.SetLinearGradient -25000, -850000, -6000, -850000, RGB(67, 66, 66), RGB(143, 143, 143), WrapModeTileFlipXY
    End With
    MyCurves.Add path

    Set path = New waPath
    With path
        .StartSubPath
        .AddLineSegment 66451, -749902
        .AddLineSegment -22846, -839200
        .AddBezierSegment -22846, -839200, -24227, -840581, -25084, -842486, -25084, -844582
        .AddLineSegment -25084, -848815
        .AddBezierSegment -25084, -848815, -25084, -853006, -21655, -856435, -17464, -856435
        .AddLineSegment -11749, -856435
        .AddLineSegment -6034, -856435
        .AddBezierSegment -6034, -856435, -3939, -856435, -2034, -855577, -652, -854197
        .AddLineSegment 88644, -764899
        .AddBezierSegment 88644, -764899, 90025, -763518, 90883, -761613, 90883, -759517
        .AddLineSegment 90883, -755284
        .AddBezierSegment 90883, -755284, 90883, -751093, 87454, -747664, 83263, -747664
        .AddLineSegment 77548, -747664
        .AddLineSegment 71833, -747664
        .AddBezierSegment 71833, -747664, 69737, -747664, 67832, -748521, 66451, -749902
        .CloseSubPath

        Set .brush = New waBrush
        .brush.SetLinearGradient -10000, -691000, 75000, -765000, RGB(164, 185, 195), RGB(218, 218, 218), WrapModeTileFlipXY
    End With
    MyCurves.Add path

    Set path = New waPath
    With path
        .StartSubPath
        .AddLineSegment -17782, -854711
        .AddLineSegment -6352, -854711
        .AddBezierSegment -6352, -854711, -3264, -854711, -732, -852178, -732, -849091
        .AddLineSegment -732, -844858
        .AddBezierSegment -732, -844858, -732, -841771, -3264, -839238, -6352, -839238
        .AddLineSegment -17782, -839238
        .AddBezierSegment -17782, -839238, -20869, -839238, -23402, -841771, -23402, -844858
        .AddLineSegment -23402, -849091
        .AddBezierSegment -23402, -849091, -23402, -852178, -20869, -854711, -17782, -854711
        .CloseSubPath

        Set .brush = New waBrush
        .brush.SetLinearGradient -20000, -835000, -4000, -846000, RGB(114, 114, 113), RGB(189, 182, 182), WrapModeTileFlipXY
    End With
    MyCurves.Add path

    Set path = New waPath
    With path
        .StartSubPath
        .AddLineSegment -1584829, -770784
        .AddLineSegment -1585537, -756279
        .AddLineSegment -1566434, -756279
        .AddLineSegment -1567069, -770784
        .AddBezierSegment -1567069, -770784, -1570695, -771828, -1581336, -772120, -1584829, -770784
        .CloseSubPath

        Set .brush = New waBrush
        .brush.SetLinearGradient -1595000, -764000, -1576000, -764000, RGB(67, 66, 66), RGB(143, 143, 143), WrapModeTileFlipXY
    End With
    MyCurves.Add path

    Set path = New waPath
    With path
        .StartSubPath
        .AddLineSegment -1669919, -856615
        .AddLineSegment -1670627, -842109
        .AddLineSegment -1651524, -842109
        .AddLineSegment -1652159, -856615
        .AddBezierSegment -1652159, -856615, -1655785, -857659, -1666427, -857951, -1669919, -856615
        .CloseSubPath

        Set .brush = New waBrush
        .brush.SetLinearGradient -1680000, -850000, -1661000, -850000, RGB(67, 66, 66), RGB(143, 143, 143), WrapModeTileFlipXY
    End With
    MyCurves.Add path

    Set path = New waPath
    With path
        .StartSubPath
        .AddLineSegment -1588515, -749902
        .AddLineSegment -1677811, -839200
        .AddBezierSegment -1677811, -839200, -1679193, -840581, -1680050, -842486, -1680050, -844582
        .AddLineSegment -1680050, -848815
        .AddBezierSegment -1680050, -848815, -1680050, -853006, -1676621, -856435, -1672430, -856435
        .AddLineSegment -1666715, -856435
        .AddLineSegment -1661000, -856435
        .AddBezierSegment -1661000, -856435, -1658904, -856435, -1656999, -855577, -1655618, -854197
        .AddLineSegment -1566321, -764899
        .AddBezierSegment -1566321, -764899, -1564940, -763518, -1564083, -761613, -1564083, -759517
        .AddLineSegment -1564083, -755284
        .AddBezierSegment -1564083, -755284, -1564083, -751093, -1567512, -747664, -1571703, -747664
        .AddLineSegment -1577418, -747664
        .AddLineSegment -1583133, -747664
        .AddBezierSegment -1583133, -747664, -1585228, -747664, -1587133, -748521, -1588515, -749902
        .CloseSubPath

        Set .brush = New waBrush
        .brush.SetLinearGradient -1664000, -691000, -1580000, -765000, RGB(164, 185, 195), RGB(218, 218, 218), WrapModeTileFlipXY
    End With
    MyCurves.Add path

    Set path = New waPath
    With path
        .StartSubPath
        .AddLineSegment 14134, -401920
        .AddLineSegment 39621, -401920
        .AddLineSegment 39621, -380859
        .AddLineSegment 14134, -380859
        .AddLineSegment 14134, -401920
        .CloseSubPath

        Set .brush = New waBrush
        .brush.SetLinearGradient 27000, -412000, 27000, -434000, RGB(43, 42, 41), RGB(141, 141, 141), WrapModeTileFlipXY
    End With
    MyCurves.Add path

    Set path = New waPath
    With path
        .StartSubPath
        .AddLineSegment -1639621, -401920
        .AddLineSegment -1614134, -401920
        .AddLineSegment -1614134, -380859
        .AddLineSegment -1639621, -380859
        .AddLineSegment -1639621, -401920
        .CloseSubPath

        Set .brush = New waBrush
        .brush.SetLinearGradient -1627000, -412000, -1627000, -434000, RGB(43, 42, 41), RGB(141, 141, 141), WrapModeTileFlipXY
    End With
    MyCurves.Add path

    Set path = New waPath
    With path
        .StartSubPath
        .AddLineSegment -1636676, -57346
        .AddLineSegment -1618050, -57346
        .AddLineSegment -1618050, -798107
        .AddLineSegment -1636676, -798107
        .AddLineSegment -1636676, -57346
        .CloseSubPath

        Set .brush = New waBrush
        .brush.SetLinearGradient -1646000, -428000, -1627000, -428000, RGB(164, 185, 195), RGB(208, 206, 210), WrapModeTileFlipXY
    End With
    MyCurves.Add path

    Set path = New waPath
    With path
        .StartSubPath
        .AddLineSegment 16050, -434570
        .AddLineSegment 16050, -463287
        .AddLineSegment -1616050, -463287
        .AddLineSegment -1616050, -434570
        .AddLineSegment 16050, -434570
        .CloseSubPath

        Set .brush = New waBrush
        .brush.SetLinearGradient -800000, -412000, -800000, -441000, RGB(164, 185, 195), RGB(214, 211, 218), WrapModeTileFlipXY
    End With
    MyCurves.Add path

    Set path = New waPath
    With path
        .StartSubPath
        .AddLineSegment 18050, -57346
        .AddLineSegment 36676, -57346
        .AddLineSegment 36676, -798107
        .AddLineSegment 18050, -798107
        .AddLineSegment 18050, -57346
        .CloseSubPath

        Set .brush = New waBrush
        .brush.SetLinearGradient 9000, -428000, 27000, -428000, RGB(164, 185, 195), RGB(208, 206, 210), WrapModeTileFlipXY
    End With
    MyCurves.Add path

    Set path = New waPath
    With path
        .StartSubPath
        .AddLineSegment 16050, -494084
        .AddLineSegment 16050, -522801
        .AddLineSegment -1616050, -522801
        .AddLineSegment -1616050, -494084
        .AddLineSegment 16050, -494084
        .CloseSubPath

        Set .brush = New waBrush
        .brush.SetLinearGradient -800000, -472000, -800000, -501000, RGB(164, 185, 195), RGB(214, 211, 218), WrapModeTileFlipXY
    End With
    MyCurves.Add path

    Set path = New waPath
    With path
        .StartSubPath
        .AddLineSegment 14134, -382007
        .AddLineSegment 14134, -400634
        .AddLineSegment -1614134, -400634
        .AddLineSegment -1614134, -382007
        .AddLineSegment 14134, -382007
        .CloseSubPath

        Set .brush = New waBrush
        .brush.SetLinearGradient -800000, -373000, -800000, -391000, RGB(164, 185, 195), RGB(208, 206, 210), WrapModeTileFlipXY
    End With
    MyCurves.Add path

    Set path = New waPath
    With path
        .StartSubPath
        .AddLineSegment -1618050, -539769
        .AddLineSegment -1616050, -539769
        .AddLineSegment -1616050, -409805
        .AddLineSegment -1618050, -409805
        .AddLineSegment -1618050, -539769
        .CloseSubPath

        Set .brush = New waBrush
        .brush.SetColor RGB(142, 164, 174), 100
    End With
    MyCurves.Add path

    Set path = New waPath
    With path
        .StartSubPath
        .AddLineSegment 16050, -539769
        .AddLineSegment 18050, -539769
        .AddLineSegment 18050, -409805
        .AddLineSegment 16050, -409805
        .AddLineSegment 16050, -539769
        .CloseSubPath

        Set .brush = New waBrush
        .brush.SetColor RGB(142, 164, 174), 100
    End With
    MyCurves.Add path

    Set path = New waPath
    With path
        .StartSubPath
        .AddLineSegment 36676, -508548
        .AddLineSegment 38676, -508548
        .AddLineSegment 38676, -409805
        .AddLineSegment 36676, -409805
        .AddLineSegment 36676, -508548
        .CloseSubPath

        Set .brush = New waBrush
        .brush.SetColor RGB(142, 164, 174), 100
    End With
    MyCurves.Add path

    Set path = New waPath
    With path
        .StartSubPath
        .AddLineSegment -1638676, -508548
        .AddLineSegment -1636676, -508548
        .AddLineSegment -1636676, -409805
        .AddLineSegment -1638676, -409805
        .AddLineSegment -1638676, -508548
        .CloseSubPath

        Set .brush = New waBrush
        .brush.SetColor RGB(142, 164, 174), 100
    End With
    MyCurves.Add path

    Set path = New waPath
    With path
        .StartSubPath
        .AddLineSegment -1672747, -854711
        .AddLineSegment -1661317, -854711
        .AddBezierSegment -1661317, -854711, -1658230, -854711, -1655697, -852178, -1655697, -849091
        .AddLineSegment -1655697, -844858
        .AddBezierSegment -1655697, -844858, -1655697, -841771, -1658230, -839238, -1661317, -839238
        .AddLineSegment -1672747, -839238
        .AddBezierSegment -1672747, -839238, -1675834, -839238, -1678367, -841771, -1678367, -844858
        .AddLineSegment -1678367, -849091
        .AddBezierSegment -1678367, -849091, -1678367, -852178, -1675834, -854711, -1672747, -854711
        .CloseSubPath

        Set .brush = New waBrush
        .brush.SetLinearGradient -1675000, -835000, -1659000, -846000, RGB(114, 114, 113), RGB(189, 182, 182), WrapModeTileFlipXY
    End With
    MyCurves.Add path

    Set path = New waPath
    With path
        .StartSubPath
        .AddLineSegment -1670499, -57346
        .AddLineSegment 70500, -57346
        .AddLineSegment 70500, -25346
        .AddLineSegment -1670499, -25346
        .AddLineSegment -1670499, -57346
        .CloseSubPath

        Set .brush = New waBrush
        .brush.SetColor RGB(164, 185, 195), 100
    End With
    MyCurves.Add path

    Set path = New waPath
    With path
        .StartSubPath
        .AddLineSegment -1670499, -17824
        .AddLineSegment 70500, -17824
        .AddLineSegment 70500, -25346
        .AddLineSegment -1670499, -25346
        .AddLineSegment -1670499, -17824
        .CloseSubPath

        Set .brush = New waBrush
        .brush.SetColor RGB(254, 254, 254), 100
    End With
    MyCurves.Add path

    Set path = New waPath
    With path
        .StartSubPath
        .AddLineSegment -645498, -25346
        .AddLineSegment -644498, -25346
        .AddLineSegment -644498, -21737
        .AddLineSegment -645498, -21737
        .AddLineSegment -645498, -25346
        .CloseSubPath

        .StartSubPath
        .AddLineSegment -5500, -25346
        .AddLineSegment -4500, -25346
        .AddLineSegment -4500, -21737
        .AddLineSegment -5500, -21737
        .AddLineSegment -5500, -25346
        .CloseSubPath

        .StartSubPath
        .AddLineSegment -500, -25346
        .AddLineSegment 500, -25346
        .AddLineSegment 500, -17824
        .AddLineSegment -500, -17824
        .AddLineSegment -500, -25346
        .CloseSubPath

        .StartSubPath
        .AddLineSegment 39500, -25346
        .AddLineSegment 40500, -25346
        .AddLineSegment 40500, -17824
        .AddLineSegment 39500, -17824
        .AddLineSegment 39500, -25346
        .CloseSubPath

        .StartSubPath
        .AddLineSegment 49500, -25346
        .AddLineSegment 50500, -25346
        .AddLineSegment 50500, -17824
        .AddLineSegment 49500, -17824
        .AddLineSegment 49500, -25346
        .CloseSubPath

        .StartSubPath
        .AddLineSegment 29500, -25346
        .AddLineSegment 30500, -25346
        .AddLineSegment 30500, -17824
        .AddLineSegment 29500, -17824
        .AddLineSegment 29500, -25346
        .CloseSubPath

        .StartSubPath
        .AddLineSegment 19500, -25346
        .AddLineSegment 20500, -25346
        .AddLineSegment 20500, -17824
        .AddLineSegment 19500, -17824
        .AddLineSegment 19500, -25346
        .CloseSubPath

        .StartSubPath
        .AddLineSegment 9500, -25346
        .AddLineSegment 10500, -25346
        .AddLineSegment 10500, -17824
        .AddLineSegment 9500, -17824
        .AddLineSegment 9500, -25346
        .CloseSubPath

        .StartSubPath
        .AddLineSegment -10500, -25346
        .AddLineSegment -9500, -25346
        .AddLineSegment -9500, -17824
        .AddLineSegment -10500, -17824
        .AddLineSegment -10500, -25346
        .CloseSubPath

        .StartSubPath
        .AddLineSegment 34500, -25346
        .AddLineSegment 35500, -25346
        .AddLineSegment 35500, -21737
        .AddLineSegment 34500, -21737
        .AddLineSegment 34500, -25346
        .CloseSubPath

        .StartSubPath
        .AddLineSegment 54500, -25346
        .AddLineSegment 55500, -25346
        .AddLineSegment 55500, -21737
        .AddLineSegment 54500, -21737
        .AddLineSegment 54500, -25346
        .CloseSubPath

        .StartSubPath
        .AddLineSegment 44500, -25346
        .AddLineSegment 45500, -25346
        .AddLineSegment 45500, -21737
        .AddLineSegment 44500, -21737
        .AddLineSegment 44500, -25346
        .CloseSubPath

        .StartSubPath
        .AddLineSegment 24500, -25346
        .AddLineSegment 25500, -25346
        .AddLineSegment 25500, -21737
        .AddLineSegment 24500, -21737
        .AddLineSegment 24500, -25346
        .CloseSubPath

        .StartSubPath
        .AddLineSegment 14500, -25346
        .AddLineSegment 15500, -25346
        .AddLineSegment 15500, -21737
        .AddLineSegment 14500, -21737
        .AddLineSegment 14500, -25346
        .CloseSubPath

        .StartSubPath
        .AddLineSegment 4500, -25346
        .AddLineSegment 5500, -25346
        .AddLineSegment 5500, -21737
        .AddLineSegment 4500, -21737
        .AddLineSegment 4500, -25346
        .CloseSubPath

        .StartSubPath
        .AddLineSegment -15500, -25346
        .AddLineSegment -14500, -25346
        .AddLineSegment -14500, -21737
        .AddLineSegment -15500, -21737
        .AddLineSegment -15500, -25346
        .CloseSubPath

        .StartSubPath
        .AddLineSegment -325499, -25346
        .AddLineSegment -324499, -25346
        .AddLineSegment -324499, -21737
        .AddLineSegment -325499, -21737
        .AddLineSegment -325499, -25346
        .CloseSubPath

        .StartSubPath
        .AddLineSegment -165499, -25346
        .AddLineSegment -164499, -25346
        .AddLineSegment -164499, -21737
        .AddLineSegment -165499, -21737
        .AddLineSegment -165499, -25346
        .CloseSubPath

        .StartSubPath
        .AddLineSegment -160499, -25346
        .AddLineSegment -159499, -25346
        .AddLineSegment -159499, -17824
        .AddLineSegment -160499, -17824
        .AddLineSegment -160499, -25346
        .CloseSubPath

        .StartSubPath
        .AddLineSegment -80500, -25346
        .AddLineSegment -79500, -25346
        .AddLineSegment -79500, -17824
        .AddLineSegment -80500, -17824
        .AddLineSegment -80500, -25346
        .CloseSubPath

        .StartSubPath
        .AddLineSegment -40500, -25346
        .AddLineSegment -39500, -25346
        .AddLineSegment -39500, -17824
        .AddLineSegment -40500, -17824
        .AddLineSegment -40500, -25346
        .CloseSubPath

        .StartSubPath
        .AddLineSegment -20500, -25346
        .AddLineSegment -19500, -25346
        .AddLineSegment -19500, -17824
        .AddLineSegment -20500, -17824
        .AddLineSegment -20500, -25346
        .CloseSubPath

        .StartSubPath
        .AddLineSegment -30500, -25346
        .AddLineSegment -29500, -25346
        .AddLineSegment -29500, -17824
        .AddLineSegment -30500, -17824
        .AddLineSegment -30500, -25346
        .CloseSubPath

        .StartSubPath
        .AddLineSegment -50500, -25346
        .AddLineSegment -49500, -25346
        .AddLineSegment -49500, -17824
        .AddLineSegment -50500, -17824
        .AddLineSegment -50500, -25346
        .CloseSubPath

        .StartSubPath
        .AddLineSegment -60500, -25346
        .AddLineSegment -59500, -25346
        .AddLineSegment -59500, -17824
        .AddLineSegment -60500, -17824
        .AddLineSegment -60500, -25346
        .CloseSubPath

        .StartSubPath
        .AddLineSegment -70500, -25346
        .AddLineSegment -69500, -25346
        .AddLineSegment -69500, -17824
        .AddLineSegment -70500, -17824
        .AddLineSegment -70500, -25346
        .CloseSubPath

        .StartSubPath
        .AddLineSegment -90500, -25346
        .AddLineSegment -89500, -25346
        .AddLineSegment -89500, -17824
        .AddLineSegment -90500, -17824
        .AddLineSegment -90500, -25346
        .CloseSubPath

        .StartSubPath
        .AddLineSegment -120499, -25346
        .AddLineSegment -119499, -25346
        .AddLineSegment -119499, -17824
        .AddLineSegment -120499, -17824
        .AddLineSegment -120499, -25346
        .CloseSubPath

        .StartSubPath
        .AddLineSegment -100500, -25346
        .AddLineSegment -99500, -25346
        .AddLineSegment -99500, -17824
        .AddLineSegment -100500, -17824
        .AddLineSegment -100500, -25346
        .CloseSubPath

        .StartSubPath
        .AddLineSegment -110499, -25346
        .AddLineSegment -109499, -25346
        .AddLineSegment -109499, -17824
        .AddLineSegment -110499, -17824
        .AddLineSegment -110499, -25346
        .CloseSubPath

        .StartSubPath
        .AddLineSegment -130499, -25346
        .AddLineSegment -129499, -25346
        .AddLineSegment -129499, -17824
        .AddLineSegment -130499, -17824
        .AddLineSegment -130499, -25346
        .CloseSubPath

        .StartSubPath
        .AddLineSegment -140500, -25346
        .AddLineSegment -139500, -25346
        .AddLineSegment -139500, -17824
        .AddLineSegment -140500, -17824
        .AddLineSegment -140500, -25346
        .CloseSubPath

        .StartSubPath
        .AddLineSegment -150499, -25346
        .AddLineSegment -149499, -25346
        .AddLineSegment -149499, -17824
        .AddLineSegment -150499, -17824
        .AddLineSegment -150499, -25346
        .CloseSubPath

        .StartSubPath
        .AddLineSegment -170499, -25346
        .AddLineSegment -169499, -25346
        .AddLineSegment -169499, -17824
        .AddLineSegment -170499, -17824
        .AddLineSegment -170499, -25346
        .CloseSubPath

        .StartSubPath
        .AddLineSegment -85500, -25346
        .AddLineSegment -84500, -25346
        .AddLineSegment -84500, -21737
        .AddLineSegment -85500, -21737
        .AddLineSegment -85500, -25346
        .CloseSubPath

        .StartSubPath
        .AddLineSegment -45500, -25346
        .AddLineSegment -44500, -25346
        .AddLineSegment -44500, -21737
        .AddLineSegment -45500, -21737
        .AddLineSegment -45500, -25346
        .CloseSubPath

        .StartSubPath
        .AddLineSegment -25500, -25346
        .AddLineSegment -24500, -25346
        .AddLineSegment -24500, -21737
        .AddLineSegment -25500, -21737
        .AddLineSegment -25500, -25346
        .CloseSubPath

        .StartSubPath
        .AddLineSegment -35500, -25346
        .AddLineSegment -34500, -25346
        .AddLineSegment -34500, -21737
        .AddLineSegment -35500, -21737
        .AddLineSegment -35500, -25346
        .CloseSubPath

        .StartSubPath
        .AddLineSegment -55500, -25346
        .AddLineSegment -54500, -25346
        .AddLineSegment -54500, -21737
        .AddLineSegment -55500, -21737
        .AddLineSegment -55500, -25346
        .CloseSubPath

        .StartSubPath
        .AddLineSegment -65500, -25346
        .AddLineSegment -64500, -25346
        .AddLineSegment -64500, -21737
        .AddLineSegment -65500, -21737
        .AddLineSegment -65500, -25346
        .CloseSubPath

        .StartSubPath
        .AddLineSegment -75500, -25346
        .AddLineSegment -74500, -25346
        .AddLineSegment -74500, -21737
        .AddLineSegment -75500, -21737
        .AddLineSegment -75500, -25346
        .CloseSubPath

        .StartSubPath
        .AddLineSegment -95500, -25346
        .AddLineSegment -94500, -25346
        .AddLineSegment -94500, -21737
        .AddLineSegment -95500, -21737
        .AddLineSegment -95500, -25346
        .CloseSubPath

        .StartSubPath
        .AddLineSegment -125499, -25346
        .AddLineSegment -124499, -25346
        .AddLineSegment -124499, -21737
        .AddLineSegment -125499, -21737
        .AddLineSegment -125499, -25346
        .CloseSubPath

        .StartSubPath
        .AddLineSegment -105500, -25346
        .AddLineSegment -104500, -25346
        .AddLineSegment -104500, -21737
        .AddLineSegment -105500, -21737
        .AddLineSegment -105500, -25346
        .CloseSubPath

        .StartSubPath
        .AddLineSegment -115499, -25346
        .AddLineSegment -114499, -25346
        .AddLineSegment -114499, -21737
        .AddLineSegment -115499, -21737
        .AddLineSegment -115499, -25346
        .CloseSubPath

        .StartSubPath
        .AddLineSegment -135499, -25346
        .AddLineSegment -134499, -25346
        .AddLineSegment -134499, -21737
        .AddLineSegment -135499, -21737
        .AddLineSegment -135499, -25346
        .CloseSubPath

        .StartSubPath
        .AddLineSegment -145500, -25346
        .AddLineSegment -144500, -25346
        .AddLineSegment -144500, -21737
        .AddLineSegment -145500, -21737
        .AddLineSegment -145500, -25346
        .CloseSubPath

        .StartSubPath
        .AddLineSegment -155499, -25346
        .AddLineSegment -154499, -25346
        .AddLineSegment -154499, -21737
        .AddLineSegment -155499, -21737
        .AddLineSegment -155499, -25346
        .CloseSubPath

        .StartSubPath
        .AddLineSegment -175499, -25346
        .AddLineSegment -174499, -25346
        .AddLineSegment -174499, -21737
        .AddLineSegment -175499, -21737
        .AddLineSegment -175499, -25346
        .CloseSubPath

        .StartSubPath
        .AddLineSegment -320499, -25346
        .AddLineSegment -319499, -25346
        .AddLineSegment -319499, -17824
        .AddLineSegment -320499, -17824
        .AddLineSegment -320499, -25346
        .CloseSubPath

        .StartSubPath
        .AddLineSegment -240499, -25346
        .AddLineSegment -239499, -25346
        .AddLineSegment -239499, -17824
        .AddLineSegment -240499, -17824
        .AddLineSegment -240499, -25346
        .CloseSubPath

        .StartSubPath
        .AddLineSegment -200499, -25346
        .AddLineSegment -199499, -25346
        .AddLineSegment -199499, -17824
        .AddLineSegment -200499, -17824
        .AddLineSegment -200499, -25346
        .CloseSubPath

        .StartSubPath
        .AddLineSegment -180499, -25346
        .AddLineSegment -179499, -25346
        .AddLineSegment -179499, -17824
        .AddLineSegment -180499, -17824
        .AddLineSegment -180499, -25346
        .CloseSubPath

        .StartSubPath
        .AddLineSegment -190499, -25346
        .AddLineSegment -189499, -25346
        .AddLineSegment -189499, -17824
        .AddLineSegment -190499, -17824
        .AddLineSegment -190499, -25346
        .CloseSubPath

        .StartSubPath
        .AddLineSegment -210499, -25346
        .AddLineSegment -209499, -25346
        .AddLineSegment -209499, -17824
        .AddLineSegment -210499, -17824
        .AddLineSegment -210499, -25346
        .CloseSubPath

        .StartSubPath
        .AddLineSegment -220499, -25346
        .AddLineSegment -219499, -25346
        .AddLineSegment -219499, -17824
        .AddLineSegment -220499, -17824
        .AddLineSegment -220499, -25346
        .CloseSubPath

        .StartSubPath
        .AddLineSegment -230499, -25346
        .AddLineSegment -229499, -25346
        .AddLineSegment -229499, -17824
        .AddLineSegment -230499, -17824
        .AddLineSegment -230499, -25346
        .CloseSubPath

        .StartSubPath
        .AddLineSegment -250499, -25346
        .AddLineSegment -249499, -25346
        .AddLineSegment -249499, -17824
        .AddLineSegment -250499, -17824
        .AddLineSegment -250499, -25346
        .CloseSubPath

        .StartSubPath
        .AddLineSegment -280499, -25346
        .AddLineSegment -279499, -25346
        .AddLineSegment -279499, -17824
        .AddLineSegment -280499, -17824
        .AddLineSegment -280499, -25346
        .CloseSubPath

        .StartSubPath
        .AddLineSegment -260499, -25346
        .AddLineSegment -259499, -25346
        .AddLineSegment -259499, -17824
        .AddLineSegment -260499, -17824
        .AddLineSegment -260499, -25346
        .CloseSubPath

        .StartSubPath
        .AddLineSegment -270499, -25346
        .AddLineSegment -269499, -25346
        .AddLineSegment -269499, -17824
        .AddLineSegment -270499, -17824
        .AddLineSegment -270499, -25346
        .CloseSubPath

        .StartSubPath
        .AddLineSegment -290499, -25346
        .AddLineSegment -289499, -25346
        .AddLineSegment -289499, -17824
        .AddLineSegment -290499, -17824
        .AddLineSegment -290499, -25346
        .CloseSubPath

        .StartSubPath
        .AddLineSegment -300499, -25346
        .AddLineSegment -299499, -25346
        .AddLineSegment -299499, -17824
        .AddLineSegment -300499, -17824
        .AddLineSegment -300499, -25346
        .CloseSubPath

        .StartSubPath
        .AddLineSegment -310499, -25346
        .AddLineSegment -309499, -25346
        .AddLineSegment -309499, -17824
        .AddLineSegment -310499, -17824
        .AddLineSegment -310499, -25346
        .CloseSubPath

        .StartSubPath
        .AddLineSegment -330499, -25346
        .AddLineSegment -329499, -25346
        .AddLineSegment -329499, -17824
        .AddLineSegment -330499, -17824
        .AddLineSegment -330499, -25346
        .CloseSubPath

        .StartSubPath
        .AddLineSegment -245499, -25346
        .AddLineSegment -244499, -25346
        .AddLineSegment -244499, -21737
        .AddLineSegment -245499, -21737
        .AddLineSegment -245499, -25346
        .CloseSubPath

        .StartSubPath
        .AddLineSegment -205499, -25346
        .AddLineSegment -204499, -25346
        .AddLineSegment -204499, -21737
        .AddLineSegment -205499, -21737
        .AddLineSegment -205499, -25346
        .CloseSubPath

        .StartSubPath
        .AddLineSegment -185499, -25346
        .AddLineSegment -184499, -25346
        .AddLineSegment -184499, -21737
        .AddLineSegment -185499, -21737
        .AddLineSegment -185499, -25346
        .CloseSubPath

        .StartSubPath
        .AddLineSegment -195499, -25346
        .AddLineSegment -194499, -25346
        .AddLineSegment -194499, -21737
        .AddLineSegment -195499, -21737
        .AddLineSegment -195499, -25346
        .CloseSubPath

        .StartSubPath
        .AddLineSegment -215499, -25346
        .AddLineSegment -214499, -25346
        .AddLineSegment -214499, -21737
        .AddLineSegment -215499, -21737
        .AddLineSegment -215499, -25346
        .CloseSubPath

        .StartSubPath
        .AddLineSegment -225499, -25346
        .AddLineSegment -224499, -25346
        .AddLineSegment -224499, -21737
        .AddLineSegment -225499, -21737
        .AddLineSegment -225499, -25346
        .CloseSubPath

        .StartSubPath
        .AddLineSegment -235499, -25346
        .AddLineSegment -234499, -25346
        .AddLineSegment -234499, -21737
        .AddLineSegment -235499, -21737
        .AddLineSegment -235499, -25346
        .CloseSubPath

        .StartSubPath
        .AddLineSegment -255499, -25346
        .AddLineSegment -254499, -25346
        .AddLineSegment -254499, -21737
        .AddLineSegment -255499, -21737
        .AddLineSegment -255499, -25346
        .CloseSubPath

        .StartSubPath
        .AddLineSegment -285499, -25346
        .AddLineSegment -284499, -25346
        .AddLineSegment -284499, -21737
        .AddLineSegment -285499, -21737
        .AddLineSegment -285499, -25346
        .CloseSubPath

        .StartSubPath
        .AddLineSegment -265499, -25346
        .AddLineSegment -264499, -25346
        .AddLineSegment -264499, -21737
        .AddLineSegment -265499, -21737
        .AddLineSegment -265499, -25346
        .CloseSubPath

        .StartSubPath
        .AddLineSegment -275499, -25346
        .AddLineSegment -274499, -25346
        .AddLineSegment -274499, -21737
        .AddLineSegment -275499, -21737
        .AddLineSegment -275499, -25346
        .CloseSubPath

        .StartSubPath
        .AddLineSegment -295499, -25346
        .AddLineSegment -294499, -25346
        .AddLineSegment -294499, -21737
        .AddLineSegment -295499, -21737
        .AddLineSegment -295499, -25346
        .CloseSubPath

        .StartSubPath
        .AddLineSegment -305499, -25346
        .AddLineSegment -304499, -25346
        .AddLineSegment -304499, -21737
        .AddLineSegment -305499, -21737
        .AddLineSegment -305499, -25346
        .CloseSubPath

        .StartSubPath
        .AddLineSegment -315499, -25346
        .AddLineSegment -314499, -25346
        .AddLineSegment -314499, -21737
        .AddLineSegment -315499, -21737
        .AddLineSegment -315499, -25346
        .CloseSubPath

        .StartSubPath
        .AddLineSegment -335499, -25346
        .AddLineSegment -334499, -25346
        .AddLineSegment -334499, -21737
        .AddLineSegment -335499, -21737
        .AddLineSegment -335499, -25346
        .CloseSubPath

        .StartSubPath
        .AddLineSegment -485498, -25346
        .AddLineSegment -484498, -25346
        .AddLineSegment -484498, -21737
        .AddLineSegment -485498, -21737
        .AddLineSegment -485498, -25346
        .CloseSubPath

        .StartSubPath
        .AddLineSegment -480498, -25346
        .AddLineSegment -479498, -25346
        .AddLineSegment -479498, -17824
        .AddLineSegment -480498, -17824
        .AddLineSegment -480498, -25346
        .CloseSubPath

        .StartSubPath
        .AddLineSegment -400498, -25346
        .AddLineSegment -399498, -25346
        .AddLineSegment -399498, -17824
        .AddLineSegment -400498, -17824
        .AddLineSegment -400498, -25346
        .CloseSubPath

        .StartSubPath
        .AddLineSegment -360499, -25346
        .AddLineSegment -359499, -25346
        .AddLineSegment -359499, -17824
        .AddLineSegment -360499, -17824
        .AddLineSegment -360499, -25346
        .CloseSubPath

        .StartSubPath
        .AddLineSegment -340499, -25346
        .AddLineSegment -339499, -25346
        .AddLineSegment -339499, -17824
        .AddLineSegment -340499, -17824
        .AddLineSegment -340499, -25346
        .CloseSubPath

        .StartSubPath
        .AddLineSegment -350499, -25346
        .AddLineSegment -349499, -25346
        .AddLineSegment -349499, -17824
        .AddLineSegment -350499, -17824
        .AddLineSegment -350499, -25346
        .CloseSubPath

        .StartSubPath
        .AddLineSegment -370498, -25346
        .AddLineSegment -369498, -25346
        .AddLineSegment -369498, -17824
        .AddLineSegment -370498, -17824
        .AddLineSegment -370498, -25346
        .CloseSubPath

        .StartSubPath
        .AddLineSegment -380499, -25346
        .AddLineSegment -379499, -25346
        .AddLineSegment -379499, -17824
        .AddLineSegment -380499, -17824
        .AddLineSegment -380499, -25346
        .CloseSubPath

        .StartSubPath
        .AddLineSegment -390498, -25346
        .AddLineSegment -389498, -25346
        .AddLineSegment -389498, -17824
        .AddLineSegment -390498, -17824
        .AddLineSegment -390498, -25346
        .CloseSubPath

        .StartSubPath
        .AddLineSegment -410498, -25346
        .AddLineSegment -409498, -25346
        .AddLineSegment -409498, -17824
        .AddLineSegment -410498, -17824
        .AddLineSegment -410498, -25346
        .CloseSubPath

        .StartSubPath
        .AddLineSegment -440498, -25346
        .AddLineSegment -439498, -25346
        .AddLineSegment -439498, -17824
        .AddLineSegment -440498, -17824
        .AddLineSegment -440498, -25346
        .CloseSubPath

        .StartSubPath
        .AddLineSegment -420498, -25346
        .AddLineSegment -419498, -25346
        .AddLineSegment -419498, -17824
        .AddLineSegment -420498, -17824
        .AddLineSegment -420498, -25346
        .CloseSubPath

        .StartSubPath
        .AddLineSegment -430498, -25346
        .AddLineSegment -429498, -25346
        .AddLineSegment -429498, -17824
        .AddLineSegment -430498, -17824
        .AddLineSegment -430498, -25346
        .CloseSubPath

        .StartSubPath
        .AddLineSegment -450498, -25346
        .AddLineSegment -449498, -25346
        .AddLineSegment -449498, -17824
        .AddLineSegment -450498, -17824
        .AddLineSegment -450498, -25346
        .CloseSubPath

        .StartSubPath
        .AddLineSegment -460498, -25346
        .AddLineSegment -459498, -25346
        .AddLineSegment -459498, -17824
        .AddLineSegment -460498, -17824
        .AddLineSegment -460498, -25346
        .CloseSubPath

        .StartSubPath
        .AddLineSegment -470498, -25346
        .AddLineSegment -469498, -25346
        .AddLineSegment -469498, -17824
        .AddLineSegment -470498, -17824
        .AddLineSegment -470498, -25346
        .CloseSubPath

        .StartSubPath
        .AddLineSegment -490498, -25346
        .AddLineSegment -489498, -25346
        .AddLineSegment -489498, -17824
        .AddLineSegment -490498, -17824
        .AddLineSegment -490498, -25346
        .CloseSubPath

        .StartSubPath
        .AddLineSegment -405498, -25346
        .AddLineSegment -404498, -25346
        .AddLineSegment -404498, -21737
        .AddLineSegment -405498, -21737
        .AddLineSegment -405498, -25346
        .CloseSubPath

        .StartSubPath
        .AddLineSegment -365499, -25346
        .AddLineSegment -364499, -25346
        .AddLineSegment -364499, -21737
        .AddLineSegment -365499, -21737
        .AddLineSegment -365499, -25346
        .CloseSubPath

        .StartSubPath
        .AddLineSegment -345499, -25346
        .AddLineSegment -344499, -25346
        .AddLineSegment -344499, -21737
        .AddLineSegment -345499, -21737
        .AddLineSegment -345499, -25346
        .CloseSubPath

        .StartSubPath
        .AddLineSegment -355499, -25346
        .AddLineSegment -354499, -25346
        .AddLineSegment -354499, -21737
        .AddLineSegment -355499, -21737
        .AddLineSegment -355499, -25346
        .CloseSubPath

        .StartSubPath
        .AddLineSegment -375498, -25346
        .AddLineSegment -374498, -25346
        .AddLineSegment -374498, -21737
        .AddLineSegment -375498, -21737
        .AddLineSegment -375498, -25346
        .CloseSubPath

        .StartSubPath
        .AddLineSegment -385499, -25346
        .AddLineSegment -384499, -25346
        .AddLineSegment -384499, -21737
        .AddLineSegment -385499, -21737
        .AddLineSegment -385499, -25346
        .CloseSubPath

        .StartSubPath
        .AddLineSegment -395498, -25346
        .AddLineSegment -394498, -25346
        .AddLineSegment -394498, -21737
        .AddLineSegment -395498, -21737
        .AddLineSegment -395498, -25346
        .CloseSubPath

        .StartSubPath
        .AddLineSegment -415498, -25346
        .AddLineSegment -414498, -25346
        .AddLineSegment -414498, -21737
        .AddLineSegment -415498, -21737
        .AddLineSegment -415498, -25346
        .CloseSubPath

        .StartSubPath
        .AddLineSegment -445498, -25346
        .AddLineSegment -444498, -25346
        .AddLineSegment -444498, -21737
        .AddLineSegment -445498, -21737
        .AddLineSegment -445498, -25346
        .CloseSubPath

        .StartSubPath
        .AddLineSegment -425498, -25346
        .AddLineSegment -424498, -25346
        .AddLineSegment -424498, -21737
        .AddLineSegment -425498, -21737
        .AddLineSegment -425498, -25346
        .CloseSubPath

        .StartSubPath
        .AddLineSegment -435498, -25346
        .AddLineSegment -434498, -25346
        .AddLineSegment -434498, -21737
        .AddLineSegment -435498, -21737
        .AddLineSegment -435498, -25346
        .CloseSubPath

        .StartSubPath
        .AddLineSegment -455498, -25346
        .AddLineSegment -454498, -25346
        .AddLineSegment -454498, -21737
        .AddLineSegment -455498, -21737
        .AddLineSegment -455498, -25346
        .CloseSubPath

        .StartSubPath
        .AddLineSegment -465498, -25346
        .AddLineSegment -464498, -25346
        .AddLineSegment -464498, -21737
        .AddLineSegment -465498, -21737
        .AddLineSegment -465498, -25346
        .CloseSubPath

        .StartSubPath
        .AddLineSegment -475498, -25346
        .AddLineSegment -474498, -25346
        .AddLineSegment -474498, -21737
        .AddLineSegment -475498, -21737
        .AddLineSegment -475498, -25346
        .CloseSubPath

        .StartSubPath
        .AddLineSegment -495498, -25346
        .AddLineSegment -494498, -25346
        .AddLineSegment -494498, -21737
        .AddLineSegment -495498, -21737
        .AddLineSegment -495498, -25346
        .CloseSubPath

        .StartSubPath
        .AddLineSegment -640498, -25346
        .AddLineSegment -639498, -25346
        .AddLineSegment -639498, -17824
        .AddLineSegment -640498, -17824
        .AddLineSegment -640498, -25346
        .CloseSubPath

        .StartSubPath
        .AddLineSegment -560498, -25346
        .AddLineSegment -559498, -25346
        .AddLineSegment -559498, -17824
        .AddLineSegment -560498, -17824
        .AddLineSegment -560498, -25346
        .CloseSubPath

        .StartSubPath
        .AddLineSegment -520498, -25346
        .AddLineSegment -519498, -25346
        .AddLineSegment -519498, -17824
        .AddLineSegment -520498, -17824
        .AddLineSegment -520498, -25346
        .CloseSubPath

        .StartSubPath
        .AddLineSegment -500498, -25346
        .AddLineSegment -499498, -25346
        .AddLineSegment -499498, -17824
        .AddLineSegment -500498, -17824
        .AddLineSegment -500498, -25346
        .CloseSubPath

        .StartSubPath
        .AddLineSegment -510498, -25346
        .AddLineSegment -509498, -25346
        .AddLineSegment -509498, -17824
        .AddLineSegment -510498, -17824
        .AddLineSegment -510498, -25346
        .CloseSubPath

        .StartSubPath
        .AddLineSegment -530498, -25346
        .AddLineSegment -529498, -25346
        .AddLineSegment -529498, -17824
        .AddLineSegment -530498, -17824
        .AddLineSegment -530498, -25346
        .CloseSubPath

        .StartSubPath
        .AddLineSegment -540498, -25346
        .AddLineSegment -539498, -25346
        .AddLineSegment -539498, -17824
        .AddLineSegment -540498, -17824
        .AddLineSegment -540498, -25346
        .CloseSubPath

        .StartSubPath
        .AddLineSegment -550498, -25346
        .AddLineSegment -549498, -25346
        .AddLineSegment -549498, -17824
        .AddLineSegment -550498, -17824
        .AddLineSegment -550498, -25346
        .CloseSubPath

        .StartSubPath
        .AddLineSegment -570498, -25346
        .AddLineSegment -569498, -25346
        .AddLineSegment -569498, -17824
        .AddLineSegment -570498, -17824
        .AddLineSegment -570498, -25346
        .CloseSubPath

        .StartSubPath
        .AddLineSegment -600498, -25346
        .AddLineSegment -599498, -25346
        .AddLineSegment -599498, -17824
        .AddLineSegment -600498, -17824
        .AddLineSegment -600498, -25346
        .CloseSubPath

        .StartSubPath
        .AddLineSegment -580498, -25346
        .AddLineSegment -579498, -25346
        .AddLineSegment -579498, -17824
        .AddLineSegment -580498, -17824
        .AddLineSegment -580498, -25346
        .CloseSubPath

        .StartSubPath
        .AddLineSegment -590498, -25346
        .AddLineSegment -589498, -25346
        .AddLineSegment -589498, -17824
        .AddLineSegment -590498, -17824
        .AddLineSegment -590498, -25346
        .CloseSubPath

        .StartSubPath
        .AddLineSegment -610498, -25346
        .AddLineSegment -609498, -25346
        .AddLineSegment -609498, -17824
        .AddLineSegment -610498, -17824
        .AddLineSegment -610498, -25346
        .CloseSubPath

        .StartSubPath
        .AddLineSegment -620498, -25346
        .AddLineSegment -619498, -25346
        .AddLineSegment -619498, -17824
        .AddLineSegment -620498, -17824
        .AddLineSegment -620498, -25346
        .CloseSubPath

        .StartSubPath
        .AddLineSegment -630498, -25346
        .AddLineSegment -629498, -25346
        .AddLineSegment -629498, -17824
        .AddLineSegment -630498, -17824
        .AddLineSegment -630498, -25346
        .CloseSubPath

        .StartSubPath
        .AddLineSegment -650497, -25346
        .AddLineSegment -649497, -25346
        .AddLineSegment -649497, -17824
        .AddLineSegment -650497, -17824
        .AddLineSegment -650497, -25346
        .CloseSubPath

        .StartSubPath
        .AddLineSegment -565498, -25346
        .AddLineSegment -564498, -25346
        .AddLineSegment -564498, -21737
        .AddLineSegment -565498, -21737
        .AddLineSegment -565498, -25346
        .CloseSubPath

        .StartSubPath
        .AddLineSegment -525498, -25346
        .AddLineSegment -524498, -25346
        .AddLineSegment -524498, -21737
        .AddLineSegment -525498, -21737
        .AddLineSegment -525498, -25346
        .CloseSubPath

        .StartSubPath
        .AddLineSegment -505498, -25346
        .AddLineSegment -504498, -25346
        .AddLineSegment -504498, -21737
        .AddLineSegment -505498, -21737
        .AddLineSegment -505498, -25346
        .CloseSubPath

        .StartSubPath
        .AddLineSegment -515498, -25346
        .AddLineSegment -514498, -25346
        .AddLineSegment -514498, -21737
        .AddLineSegment -515498, -21737
        .AddLineSegment -515498, -25346
        .CloseSubPath

        .StartSubPath
        .AddLineSegment -535498, -25346
        .AddLineSegment -534498, -25346
        .AddLineSegment -534498, -21737
        .AddLineSegment -535498, -21737
        .AddLineSegment -535498, -25346
        .CloseSubPath

        .StartSubPath
        .AddLineSegment -545498, -25346
        .AddLineSegment -544498, -25346
        .AddLineSegment -544498, -21737
        .AddLineSegment -545498, -21737
        .AddLineSegment -545498, -25346
        .CloseSubPath

        .StartSubPath
        .AddLineSegment -555498, -25346
        .AddLineSegment -554498, -25346
        .AddLineSegment -554498, -21737
        .AddLineSegment -555498, -21737
        .AddLineSegment -555498, -25346
        .CloseSubPath

        .StartSubPath
        .AddLineSegment -575498, -25346
        .AddLineSegment -574498, -25346
        .AddLineSegment -574498, -21737
        .AddLineSegment -575498, -21737
        .AddLineSegment -575498, -25346
        .CloseSubPath

        .StartSubPath
        .AddLineSegment -605498, -25346
        .AddLineSegment -604498, -25346
        .AddLineSegment -604498, -21737
        .AddLineSegment -605498, -21737
        .AddLineSegment -605498, -25346
        .CloseSubPath

        .StartSubPath
        .AddLineSegment -585498, -25346
        .AddLineSegment -584498, -25346
        .AddLineSegment -584498, -21737
        .AddLineSegment -585498, -21737
        .AddLineSegment -585498, -25346
        .CloseSubPath

        .StartSubPath
        .AddLineSegment -595498, -25346
        .AddLineSegment -594498, -25346
        .AddLineSegment -594498, -21737
        .AddLineSegment -595498, -21737
        .AddLineSegment -595498, -25346
        .CloseSubPath

        .StartSubPath
        .AddLineSegment -615498, -25346
        .AddLineSegment -614498, -25346
        .AddLineSegment -614498, -21737
        .AddLineSegment -615498, -21737
        .AddLineSegment -615498, -25346
        .CloseSubPath

        .StartSubPath
        .AddLineSegment -625498, -25346
        .AddLineSegment -624498, -25346
        .AddLineSegment -624498, -21737
        .AddLineSegment -625498, -21737
        .AddLineSegment -625498, -25346
        .CloseSubPath

        .StartSubPath
        .AddLineSegment -635498, -25346
        .AddLineSegment -634498, -25346
        .AddLineSegment -634498, -21737
        .AddLineSegment -635498, -21737
        .AddLineSegment -635498, -25346
        .CloseSubPath

        .StartSubPath
        .AddLineSegment -655497, -25346
        .AddLineSegment -654497, -25346
        .AddLineSegment -654497, -21737
        .AddLineSegment -655497, -21737
        .AddLineSegment -655497, -25346
        .CloseSubPath

        .StartSubPath
        .AddLineSegment -665497, -25346
        .AddLineSegment -664497, -25346
        .AddLineSegment -664497, -21737
        .AddLineSegment -665497, -21737
        .AddLineSegment -665497, -25346
        .CloseSubPath

        .StartSubPath
        .AddLineSegment -660497, -25346
        .AddLineSegment -659497, -25346
        .AddLineSegment -659497, -17824
        .AddLineSegment -660497, -17824
        .AddLineSegment -660497, -25346
        .CloseSubPath

        .StartSubPath
        .AddLineSegment 64500, -25346
        .AddLineSegment 65500, -25346
        .AddLineSegment 65500, -21737
        .AddLineSegment 64500, -21737
        .AddLineSegment 64500, -25346
        .CloseSubPath

        .StartSubPath
        .AddLineSegment 59500, -25346
        .AddLineSegment 60500, -25346
        .AddLineSegment 60500, -17824
        .AddLineSegment 59500, -17824
        .AddLineSegment 59500, -25346
        .CloseSubPath

        .StartSubPath
        .AddLineSegment -1375495, -25346
        .AddLineSegment -1374495, -25346
        .AddLineSegment -1374495, -21737
        .AddLineSegment -1375495, -21737
        .AddLineSegment -1375495, -25346
        .CloseSubPath

        .StartSubPath
        .AddLineSegment -735497, -25346
        .AddLineSegment -734497, -25346
        .AddLineSegment -734497, -21737
        .AddLineSegment -735497, -21737
        .AddLineSegment -735497, -25346
        .CloseSubPath

        .StartSubPath
        .AddLineSegment -730497, -25346
        .AddLineSegment -729497, -25346
        .AddLineSegment -729497, -17824
        .AddLineSegment -730497, -17824
        .AddLineSegment -730497, -25346
        .CloseSubPath

        .StartSubPath
        .AddLineSegment -690497, -25346
        .AddLineSegment -689497, -25346
        .AddLineSegment -689497, -17824
        .AddLineSegment -690497, -17824
        .AddLineSegment -690497, -25346
        .CloseSubPath

        .StartSubPath
        .AddLineSegment -680497, -25346
        .AddLineSegment -679497, -25346
        .AddLineSegment -679497, -17824
        .AddLineSegment -680497, -17824
        .AddLineSegment -680497, -25346
        .CloseSubPath

        .StartSubPath
        .AddLineSegment -700497, -25346
        .AddLineSegment -699497, -25346
        .AddLineSegment -699497, -17824
        .AddLineSegment -700497, -17824
        .AddLineSegment -700497, -25346
        .CloseSubPath

        .StartSubPath
        .AddLineSegment -710497, -25346
        .AddLineSegment -709497, -25346
        .AddLineSegment -709497, -17824
        .AddLineSegment -710497, -17824
        .AddLineSegment -710497, -25346
        .CloseSubPath

        .StartSubPath
        .AddLineSegment -720497, -25346
        .AddLineSegment -719497, -25346
        .AddLineSegment -719497, -17824
        .AddLineSegment -720497, -17824
        .AddLineSegment -720497, -25346
        .CloseSubPath

        .StartSubPath
        .AddLineSegment -740497, -25346
        .AddLineSegment -739497, -25346
        .AddLineSegment -739497, -17824
        .AddLineSegment -740497, -17824
        .AddLineSegment -740497, -25346
        .CloseSubPath

        .StartSubPath
        .AddLineSegment -695497, -25346
        .AddLineSegment -694497, -25346
        .AddLineSegment -694497, -21737
        .AddLineSegment -695497, -21737
        .AddLineSegment -695497, -25346
        .CloseSubPath

        .StartSubPath
        .AddLineSegment -675497, -25346
        .AddLineSegment -674497, -25346
        .AddLineSegment -674497, -21737
        .AddLineSegment -675497, -21737
        .AddLineSegment -675497, -25346
        .CloseSubPath

        .StartSubPath
        .AddLineSegment -685497, -25346
        .AddLineSegment -684497, -25346
        .AddLineSegment -684497, -21737
        .AddLineSegment -685497, -21737
        .AddLineSegment -685497, -25346
        .CloseSubPath

        .StartSubPath
        .AddLineSegment -705497, -25346
        .AddLineSegment -704497, -25346
        .AddLineSegment -704497, -21737
        .AddLineSegment -705497, -21737
        .AddLineSegment -705497, -25346
        .CloseSubPath

        .StartSubPath
        .AddLineSegment -715497, -25346
        .AddLineSegment -714497, -25346
        .AddLineSegment -714497, -21737
        .AddLineSegment -715497, -21737
        .AddLineSegment -715497, -25346
        .CloseSubPath

        .StartSubPath
        .AddLineSegment -725497, -25346
        .AddLineSegment -724497, -25346
        .AddLineSegment -724497, -21737
        .AddLineSegment -725497, -21737
        .AddLineSegment -725497, -25346
        .CloseSubPath

        .StartSubPath
        .AddLineSegment -745497, -25346
        .AddLineSegment -744497, -25346
        .AddLineSegment -744497, -21737
        .AddLineSegment -745497, -21737
        .AddLineSegment -745497, -25346
        .CloseSubPath

        .StartSubPath
        .AddLineSegment -1055496, -25346
        .AddLineSegment -1054496, -25346
        .AddLineSegment -1054496, -21737
        .AddLineSegment -1055496, -21737
        .AddLineSegment -1055496, -25346
        .CloseSubPath

        .StartSubPath
        .AddLineSegment -895497, -25346
        .AddLineSegment -894497, -25346
        .AddLineSegment -894497, -21737
        .AddLineSegment -895497, -21737
        .AddLineSegment -895497, -25346
        .CloseSubPath

        .StartSubPath
        .AddLineSegment -890497, -25346
        .AddLineSegment -889497, -25346
        .AddLineSegment -889497, -17824
        .AddLineSegment -890497, -17824
        .AddLineSegment -890497, -25346
        .CloseSubPath

        .StartSubPath
        .AddLineSegment -810497, -25346
        .AddLineSegment -809497, -25346
        .AddLineSegment -809497, -17824
        .AddLineSegment -810497, -17824
        .AddLineSegment -810497, -25346
        .CloseSubPath

        .StartSubPath
        .AddLineSegment -770497, -25346
        .AddLineSegment -769497, -25346
        .AddLineSegment -769497, -17824
        .AddLineSegment -770497, -17824
        .AddLineSegment -770497, -25346
        .CloseSubPath

        .StartSubPath
        .AddLineSegment -750497, -25346
        .AddLineSegment -749497, -25346
        .AddLineSegment -749497, -17824
        .AddLineSegment -750497, -17824
        .AddLineSegment -750497, -25346
        .CloseSubPath

        .StartSubPath
        .AddLineSegment -760497, -25346
        .AddLineSegment -759497, -25346
        .AddLineSegment -759497, -17824
        .AddLineSegment -760497, -17824
        .AddLineSegment -760497, -25346
        .CloseSubPath

        .StartSubPath
        .AddLineSegment -780497, -25346
        .AddLineSegment -779497, -25346
        .AddLineSegment -779497, -17824
        .AddLineSegment -780497, -17824
        .AddLineSegment -780497, -25346
        .CloseSubPath

        .StartSubPath
        .AddLineSegment -790497, -25346
        .AddLineSegment -789497, -25346
        .AddLineSegment -789497, -17824
        .AddLineSegment -790497, -17824
        .AddLineSegment -790497, -25346
        .CloseSubPath

        .StartSubPath
        .AddLineSegment -800497, -25346
        .AddLineSegment -799497, -25346
        .AddLineSegment -799497, -17824
        .AddLineSegment -800497, -17824
        .AddLineSegment -800497, -25346
        .CloseSubPath

        .StartSubPath
        .AddLineSegment -820497, -25346
        .AddLineSegment -819497, -25346
        .AddLineSegment -819497, -17824
        .AddLineSegment -820497, -17824
        .AddLineSegment -820497, -25346
        .CloseSubPath

        .StartSubPath
        .AddLineSegment -850497, -25346
        .AddLineSegment -849497, -25346
        .AddLineSegment -849497, -17824
        .AddLineSegment -850497, -17824
        .AddLineSegment -850497, -25346
        .CloseSubPath

        .StartSubPath
        .AddLineSegment -830497, -25346
        .AddLineSegment -829497, -25346
        .AddLineSegment -829497, -17824
        .AddLineSegment -830497, -17824
        .AddLineSegment -830497, -25346
        .CloseSubPath

        .StartSubPath
        .AddLineSegment -840497, -25346
        .AddLineSegment -839497, -25346
        .AddLineSegment -839497, -17824
        .AddLineSegment -840497, -17824
        .AddLineSegment -840497, -25346
        .CloseSubPath

        .StartSubPath
        .AddLineSegment -860497, -25346
        .AddLineSegment -859497, -25346
        .AddLineSegment -859497, -17824
        .AddLineSegment -860497, -17824
        .AddLineSegment -860497, -25346
        .CloseSubPath

        .StartSubPath
        .AddLineSegment -870497, -25346
        .AddLineSegment -869497, -25346
        .AddLineSegment -869497, -17824
        .AddLineSegment -870497, -17824
        .AddLineSegment -870497, -25346
        .CloseSubPath

        .StartSubPath
        .AddLineSegment -880497, -25346
        .AddLineSegment -879497, -25346
        .AddLineSegment -879497, -17824
        .AddLineSegment -880497, -17824
        .AddLineSegment -880497, -25346
        .CloseSubPath

        .StartSubPath
        .AddLineSegment -900496, -25346
        .AddLineSegment -899496, -25346
        .AddLineSegment -899496, -17824
        .AddLineSegment -900496, -17824
        .AddLineSegment -900496, -25346
        .CloseSubPath

        .StartSubPath
        .AddLineSegment -815497, -25346
        .AddLineSegment -814497, -25346
        .AddLineSegment -814497, -21737
        .AddLineSegment -815497, -21737
        .AddLineSegment -815497, -25346
        .CloseSubPath

        .StartSubPath
        .AddLineSegment -775497, -25346
        .AddLineSegment -774497, -25346
        .AddLineSegment -774497, -21737
        .AddLineSegment -775497, -21737
        .AddLineSegment -775497, -25346
        .CloseSubPath

        .StartSubPath
        .AddLineSegment -755497, -25346
        .AddLineSegment -754497, -25346
        .AddLineSegment -754497, -21737
        .AddLineSegment -755497, -21737
        .AddLineSegment -755497, -25346
        .CloseSubPath

        .StartSubPath
        .AddLineSegment -765497, -25346
        .AddLineSegment -764497, -25346
        .AddLineSegment -764497, -21737
        .AddLineSegment -765497, -21737
        .AddLineSegment -765497, -25346
        .CloseSubPath

        .StartSubPath
        .AddLineSegment -785497, -25346
        .AddLineSegment -784497, -25346
        .AddLineSegment -784497, -21737
        .AddLineSegment -785497, -21737
        .AddLineSegment -785497, -25346
        .CloseSubPath

        .StartSubPath
        .AddLineSegment -795497, -25346
        .AddLineSegment -794497, -25346
        .AddLineSegment -794497, -21737
        .AddLineSegment -795497, -21737
        .AddLineSegment -795497, -25346
        .CloseSubPath

        .StartSubPath
        .AddLineSegment -805497, -25346
        .AddLineSegment -804497, -25346
        .AddLineSegment -804497, -21737
        .AddLineSegment -805497, -21737
        .AddLineSegment -805497, -25346
        .CloseSubPath

        .StartSubPath
        .AddLineSegment -825497, -25346
        .AddLineSegment -824497, -25346
        .AddLineSegment -824497, -21737
        .AddLineSegment -825497, -21737
        .AddLineSegment -825497, -25346
        .CloseSubPath

        .StartSubPath
        .AddLineSegment -855497, -25346
        .AddLineSegment -854497, -25346
        .AddLineSegment -854497, -21737
        .AddLineSegment -855497, -21737
        .AddLineSegment -855497, -25346
        .CloseSubPath

        .StartSubPath
        .AddLineSegment -835497, -25346
        .AddLineSegment -834497, -25346
        .AddLineSegment -834497, -21737
        .AddLineSegment -835497, -21737
        .AddLineSegment -835497, -25346
        .CloseSubPath

        .StartSubPath
        .AddLineSegment -845497, -25346
        .AddLineSegment -844497, -25346
        .AddLineSegment -844497, -21737
        .AddLineSegment -845497, -21737
        .AddLineSegment -845497, -25346
        .CloseSubPath

        .StartSubPath
        .AddLineSegment -865497, -25346
        .AddLineSegment -864497, -25346
        .AddLineSegment -864497, -21737
        .AddLineSegment -865497, -21737
        .AddLineSegment -865497, -25346
        .CloseSubPath

        .StartSubPath
        .AddLineSegment -875497, -25346
        .AddLineSegment -874497, -25346
        .AddLineSegment -874497, -21737
        .AddLineSegment -875497, -21737
        .AddLineSegment -875497, -25346
        .CloseSubPath

        .StartSubPath
        .AddLineSegment -885497, -25346
        .AddLineSegment -884497, -25346
        .AddLineSegment -884497, -21737
        .AddLineSegment -885497, -21737
        .AddLineSegment -885497, -25346
        .CloseSubPath

        .StartSubPath
        .AddLineSegment -905496, -25346
        .AddLineSegment -904496, -25346
        .AddLineSegment -904496, -21737
        .AddLineSegment -905496, -21737
        .AddLineSegment -905496, -25346
        .CloseSubPath

        .StartSubPath
        .AddLineSegment -1050496, -25346
        .AddLineSegment -1049496, -25346
        .AddLineSegment -1049496, -17824
        .AddLineSegment -1050496, -17824
        .AddLineSegment -1050496, -25346
        .CloseSubPath

        .StartSubPath
        .AddLineSegment -970496, -25346
        .AddLineSegment -969496, -25346
        .AddLineSegment -969496, -17824
        .AddLineSegment -970496, -17824
        .AddLineSegment -970496, -25346
        .CloseSubPath

        .StartSubPath
        .AddLineSegment -930496, -25346
        .AddLineSegment -929496, -25346
        .AddLineSegment -929496, -17824
        .AddLineSegment -930496, -17824
        .AddLineSegment -930496, -25346
        .CloseSubPath

        .StartSubPath
        .AddLineSegment -910496, -25346
        .AddLineSegment -909496, -25346
        .AddLineSegment -909496, -17824
        .AddLineSegment -910496, -17824
        .AddLineSegment -910496, -25346
        .CloseSubPath

        .StartSubPath
        .AddLineSegment -920496, -25346
        .AddLineSegment -919496, -25346
        .AddLineSegment -919496, -17824
        .AddLineSegment -920496, -17824
        .AddLineSegment -920496, -25346
        .CloseSubPath

        .StartSubPath
        .AddLineSegment -940496, -25346
        .AddLineSegment -939496, -25346
        .AddLineSegment -939496, -17824
        .AddLineSegment -940496, -17824
        .AddLineSegment -940496, -25346
        .CloseSubPath

        .StartSubPath
        .AddLineSegment -950496, -25346
        .AddLineSegment -949496, -25346
        .AddLineSegment -949496, -17824
        .AddLineSegment -950496, -17824
        .AddLineSegment -950496, -25346
        .CloseSubPath

        .StartSubPath
        .AddLineSegment -960496, -25346
        .AddLineSegment -959496, -25346
        .AddLineSegment -959496, -17824
        .AddLineSegment -960496, -17824
        .AddLineSegment -960496, -25346
        .CloseSubPath

        .StartSubPath
        .AddLineSegment -980496, -25346
        .AddLineSegment -979496, -25346
        .AddLineSegment -979496, -17824
        .AddLineSegment -980496, -17824
        .AddLineSegment -980496, -25346
        .CloseSubPath

        .StartSubPath
        .AddLineSegment -1010496, -25346
        .AddLineSegment -1009496, -25346
        .AddLineSegment -1009496, -17824
        .AddLineSegment -1010496, -17824
        .AddLineSegment -1010496, -25346
        .CloseSubPath

        .StartSubPath
        .AddLineSegment -990496, -25346
        .AddLineSegment -989496, -25346
        .AddLineSegment -989496, -17824
        .AddLineSegment -990496, -17824
        .AddLineSegment -990496, -25346
        .CloseSubPath

        .StartSubPath
        .AddLineSegment -1000496, -25346
        .AddLineSegment -999496, -25346
        .AddLineSegment -999496, -17824
        .AddLineSegment -1000496, -17824
        .AddLineSegment -1000496, -25346
        .CloseSubPath

        .StartSubPath
        .AddLineSegment -1020496, -25346
        .AddLineSegment -1019496, -25346
        .AddLineSegment -1019496, -17824
        .AddLineSegment -1020496, -17824
        .AddLineSegment -1020496, -25346
        .CloseSubPath

        .StartSubPath
        .AddLineSegment -1030496, -25346
        .AddLineSegment -1029496, -25346
        .AddLineSegment -1029496, -17824
        .AddLineSegment -1030496, -17824
        .AddLineSegment -1030496, -25346
        .CloseSubPath

        .StartSubPath
        .AddLineSegment -1040496, -25346
        .AddLineSegment -1039496, -25346
        .AddLineSegment -1039496, -17824
        .AddLineSegment -1040496, -17824
        .AddLineSegment -1040496, -25346
        .CloseSubPath

        .StartSubPath
        .AddLineSegment -1060496, -25346
        .AddLineSegment -1059496, -25346
        .AddLineSegment -1059496, -17824
        .AddLineSegment -1060496, -17824
        .AddLineSegment -1060496, -25346
        .CloseSubPath

        .StartSubPath
        .AddLineSegment -975496, -25346
        .AddLineSegment -974496, -25346
        .AddLineSegment -974496, -21737
        .AddLineSegment -975496, -21737
        .AddLineSegment -975496, -25346
        .CloseSubPath

        .StartSubPath
        .AddLineSegment -935496, -25346
        .AddLineSegment -934496, -25346
        .AddLineSegment -934496, -21737
        .AddLineSegment -935496, -21737
        .AddLineSegment -935496, -25346
        .CloseSubPath

        .StartSubPath
        .AddLineSegment -915496, -25346
        .AddLineSegment -914496, -25346
        .AddLineSegment -914496, -21737
        .AddLineSegment -915496, -21737
        .AddLineSegment -915496, -25346
        .CloseSubPath

        .StartSubPath
        .AddLineSegment -925496, -25346
        .AddLineSegment -924496, -25346
        .AddLineSegment -924496, -21737
        .AddLineSegment -925496, -21737
        .AddLineSegment -925496, -25346
        .CloseSubPath

        .StartSubPath
        .AddLineSegment -945496, -25346
        .AddLineSegment -944496, -25346
        .AddLineSegment -944496, -21737
        .AddLineSegment -945496, -21737
        .AddLineSegment -945496, -25346
        .CloseSubPath

        .StartSubPath
        .AddLineSegment -955496, -25346
        .AddLineSegment -954496, -25346
        .AddLineSegment -954496, -21737
        .AddLineSegment -955496, -21737
        .AddLineSegment -955496, -25346
        .CloseSubPath

        .StartSubPath
        .AddLineSegment -965496, -25346
        .AddLineSegment -964496, -25346
        .AddLineSegment -964496, -21737
        .AddLineSegment -965496, -21737
        .AddLineSegment -965496, -25346
        .CloseSubPath

        .StartSubPath
        .AddLineSegment -985496, -25346
        .AddLineSegment -984496, -25346
        .AddLineSegment -984496, -21737
        .AddLineSegment -985496, -21737
        .AddLineSegment -985496, -25346
        .CloseSubPath

        .StartSubPath
        .AddLineSegment -1015496, -25346
        .AddLineSegment -1014496, -25346
        .AddLineSegment -1014496, -21737
        .AddLineSegment -1015496, -21737
        .AddLineSegment -1015496, -25346
        .CloseSubPath

        .StartSubPath
        .AddLineSegment -995496, -25346
        .AddLineSegment -994496, -25346
        .AddLineSegment -994496, -21737
        .AddLineSegment -995496, -21737
        .AddLineSegment -995496, -25346
        .CloseSubPath

        .StartSubPath
        .AddLineSegment -1005496, -25346
        .AddLineSegment -1004496, -25346
        .AddLineSegment -1004496, -21737
        .AddLineSegment -1005496, -21737
        .AddLineSegment -1005496, -25346
        .CloseSubPath

        .StartSubPath
        .AddLineSegment -1025496, -25346
        .AddLineSegment -1024496, -25346
        .AddLineSegment -1024496, -21737
        .AddLineSegment -1025496, -21737
        .AddLineSegment -1025496, -25346
        .CloseSubPath

        .StartSubPath
        .AddLineSegment -1035496, -25346
        .AddLineSegment -1034496, -25346
        .AddLineSegment -1034496, -21737
        .AddLineSegment -1035496, -21737
        .AddLineSegment -1035496, -25346
        .CloseSubPath

        .StartSubPath
        .AddLineSegment -1045496, -25346
        .AddLineSegment -1044496, -25346
        .AddLineSegment -1044496, -21737
        .AddLineSegment -1045496, -21737
        .AddLineSegment -1045496, -25346
        .CloseSubPath

        .StartSubPath
        .AddLineSegment -1065496, -25346
        .AddLineSegment -1064496, -25346
        .AddLineSegment -1064496, -21737
        .AddLineSegment -1065496, -21737
        .AddLineSegment -1065496, -25346
        .CloseSubPath

        .StartSubPath
        .AddLineSegment -1215495, -25346
        .AddLineSegment -1214495, -25346
        .AddLineSegment -1214495, -21737
        .AddLineSegment -1215495, -21737
        .AddLineSegment -1215495, -25346
        .CloseSubPath

        .StartSubPath
        .AddLineSegment -1210495, -25346
        .AddLineSegment -1209495, -25346
        .AddLineSegment -1209495, -17824
        .AddLineSegment -1210495, -17824
        .AddLineSegment -1210495, -25346
        .CloseSubPath

        .StartSubPath
        .AddLineSegment -1130496, -25346
        .AddLineSegment -1129496, -25346
        .AddLineSegment -1129496, -17824
        .AddLineSegment -1130496, -17824
        .AddLineSegment -1130496, -25346
        .CloseSubPath

        .StartSubPath
        .AddLineSegment -1090496, -25346
        .AddLineSegment -1089496, -25346
        .AddLineSegment -1089496, -17824
        .AddLineSegment -1090496, -17824
        .AddLineSegment -1090496, -25346
        .CloseSubPath

        .StartSubPath
        .AddLineSegment -1070496, -25346
        .AddLineSegment -1069496, -25346
        .AddLineSegment -1069496, -17824
        .AddLineSegment -1070496, -17824
        .AddLineSegment -1070496, -25346
        .CloseSubPath

        .StartSubPath
        .AddLineSegment -1080496, -25346
        .AddLineSegment -1079496, -25346
        .AddLineSegment -1079496, -17824
        .AddLineSegment -1080496, -17824
        .AddLineSegment -1080496, -25346
        .CloseSubPath

        .StartSubPath
        .AddLineSegment -1100496, -25346
        .AddLineSegment -1099496, -25346
        .AddLineSegment -1099496, -17824
        .AddLineSegment -1100496, -17824
        .AddLineSegment -1100496, -25346
        .CloseSubPath

        .StartSubPath
        .AddLineSegment -1110496, -25346
        .AddLineSegment -1109496, -25346
        .AddLineSegment -1109496, -17824
        .AddLineSegment -1110496, -17824
        .AddLineSegment -1110496, -25346
        .CloseSubPath

        .StartSubPath
        .AddLineSegment -1120496, -25346
        .AddLineSegment -1119496, -25346
        .AddLineSegment -1119496, -17824
        .AddLineSegment -1120496, -17824
        .AddLineSegment -1120496, -25346
        .CloseSubPath

        .StartSubPath
        .AddLineSegment -1140496, -25346
        .AddLineSegment -1139496, -25346
        .AddLineSegment -1139496, -17824
        .AddLineSegment -1140496, -17824
        .AddLineSegment -1140496, -25346
        .CloseSubPath

        .StartSubPath
        .AddLineSegment -1170495, -25346
        .AddLineSegment -1169495, -25346
        .AddLineSegment -1169495, -17824
        .AddLineSegment -1170495, -17824
        .AddLineSegment -1170495, -25346
        .CloseSubPath

        .StartSubPath
        .AddLineSegment -1150496, -25346
        .AddLineSegment -1149496, -25346
        .AddLineSegment -1149496, -17824
        .AddLineSegment -1150496, -17824
        .AddLineSegment -1150496, -25346
        .CloseSubPath

        .StartSubPath
        .AddLineSegment -1160495, -25346
        .AddLineSegment -1159495, -25346
        .AddLineSegment -1159495, -17824
        .AddLineSegment -1160495, -17824
        .AddLineSegment -1160495, -25346
        .CloseSubPath

        .StartSubPath
        .AddLineSegment -1180495, -25346
        .AddLineSegment -1179495, -25346
        .AddLineSegment -1179495, -17824
        .AddLineSegment -1180495, -17824
        .AddLineSegment -1180495, -25346
        .CloseSubPath

        .StartSubPath
        .AddLineSegment -1190495, -25346
        .AddLineSegment -1189495, -25346
        .AddLineSegment -1189495, -17824
        .AddLineSegment -1190495, -17824
        .AddLineSegment -1190495, -25346
        .CloseSubPath

        .StartSubPath
        .AddLineSegment -1200495, -25346
        .AddLineSegment -1199495, -25346
        .AddLineSegment -1199495, -17824
        .AddLineSegment -1200495, -17824
        .AddLineSegment -1200495, -25346
        .CloseSubPath

        .StartSubPath
        .AddLineSegment -1220495, -25346
        .AddLineSegment -1219495, -25346
        .AddLineSegment -1219495, -17824
        .AddLineSegment -1220495, -17824
        .AddLineSegment -1220495, -25346
        .CloseSubPath

        .StartSubPath
        .AddLineSegment -1135496, -25346
        .AddLineSegment -1134496, -25346
        .AddLineSegment -1134496, -21737
        .AddLineSegment -1135496, -21737
        .AddLineSegment -1135496, -25346
        .CloseSubPath

        .StartSubPath
        .AddLineSegment -1095496, -25346
        .AddLineSegment -1094496, -25346
        .AddLineSegment -1094496, -21737
        .AddLineSegment -1095496, -21737
        .AddLineSegment -1095496, -25346
        .CloseSubPath

        .StartSubPath
        .AddLineSegment -1075496, -25346
        .AddLineSegment -1074496, -25346
        .AddLineSegment -1074496, -21737
        .AddLineSegment -1075496, -21737
        .AddLineSegment -1075496, -25346
        .CloseSubPath

        .StartSubPath
        .AddLineSegment -1085496, -25346
        .AddLineSegment -1084496, -25346
        .AddLineSegment -1084496, -21737
        .AddLineSegment -1085496, -21737
        .AddLineSegment -1085496, -25346
        .CloseSubPath

        .StartSubPath
        .AddLineSegment -1105496, -25346
        .AddLineSegment -1104496, -25346
        .AddLineSegment -1104496, -21737
        .AddLineSegment -1105496, -21737
        .AddLineSegment -1105496, -25346
        .CloseSubPath

        .StartSubPath
        .AddLineSegment -1115496, -25346
        .AddLineSegment -1114496, -25346
        .AddLineSegment -1114496, -21737
        .AddLineSegment -1115496, -21737
        .AddLineSegment -1115496, -25346
        .CloseSubPath

        .StartSubPath
        .AddLineSegment -1125496, -25346
        .AddLineSegment -1124496, -25346
        .AddLineSegment -1124496, -21737
        .AddLineSegment -1125496, -21737
        .AddLineSegment -1125496, -25346
        .CloseSubPath

        .StartSubPath
        .AddLineSegment -1145496, -25346
        .AddLineSegment -1144496, -25346
        .AddLineSegment -1144496, -21737
        .AddLineSegment -1145496, -21737
        .AddLineSegment -1145496, -25346
        .CloseSubPath

        .StartSubPath
        .AddLineSegment -1175495, -25346
        .AddLineSegment -1174495, -25346
        .AddLineSegment -1174495, -21737
        .AddLineSegment -1175495, -21737
        .AddLineSegment -1175495, -25346
        .CloseSubPath

        .StartSubPath
        .AddLineSegment -1155496, -25346
        .AddLineSegment -1154496, -25346
        .AddLineSegment -1154496, -21737
        .AddLineSegment -1155496, -21737
        .AddLineSegment -1155496, -25346
        .CloseSubPath

        .StartSubPath
        .AddLineSegment -1165495, -25346
        .AddLineSegment -1164495, -25346
        .AddLineSegment -1164495, -21737
        .AddLineSegment -1165495, -21737
        .AddLineSegment -1165495, -25346
        .CloseSubPath

        .StartSubPath
        .AddLineSegment -1185495, -25346
        .AddLineSegment -1184495, -25346
        .AddLineSegment -1184495, -21737
        .AddLineSegment -1185495, -21737
        .AddLineSegment -1185495, -25346
        .CloseSubPath

        .StartSubPath
        .AddLineSegment -1195495, -25346
        .AddLineSegment -1194495, -25346
        .AddLineSegment -1194495, -21737
        .AddLineSegment -1195495, -21737
        .AddLineSegment -1195495, -25346
        .CloseSubPath

        .StartSubPath
        .AddLineSegment -1205495, -25346
        .AddLineSegment -1204495, -25346
        .AddLineSegment -1204495, -21737
        .AddLineSegment -1205495, -21737
        .AddLineSegment -1205495, -25346
        .CloseSubPath

        .StartSubPath
        .AddLineSegment -1225495, -25346
        .AddLineSegment -1224495, -25346
        .AddLineSegment -1224495, -21737
        .AddLineSegment -1225495, -21737
        .AddLineSegment -1225495, -25346
        .CloseSubPath

        .StartSubPath
        .AddLineSegment -1370495, -25346
        .AddLineSegment -1369495, -25346
        .AddLineSegment -1369495, -17824
        .AddLineSegment -1370495, -17824
        .AddLineSegment -1370495, -25346
        .CloseSubPath

        .StartSubPath
        .AddLineSegment -1290495, -25346
        .AddLineSegment -1289495, -25346
        .AddLineSegment -1289495, -17824
        .AddLineSegment -1290495, -17824
        .AddLineSegment -1290495, -25346
        .CloseSubPath

        .StartSubPath
        .AddLineSegment -1250495, -25346
        .AddLineSegment -1249495, -25346
        .AddLineSegment -1249495, -17824
        .AddLineSegment -1250495, -17824
        .AddLineSegment -1250495, -25346
        .CloseSubPath

        .StartSubPath
        .AddLineSegment -1230495, -25346
        .AddLineSegment -1229495, -25346
        .AddLineSegment -1229495, -17824
        .AddLineSegment -1230495, -17824
        .AddLineSegment -1230495, -25346
        .CloseSubPath

        .StartSubPath
        .AddLineSegment -1240495, -25346
        .AddLineSegment -1239495, -25346
        .AddLineSegment -1239495, -17824
        .AddLineSegment -1240495, -17824
        .AddLineSegment -1240495, -25346
        .CloseSubPath

        .StartSubPath
        .AddLineSegment -1260495, -25346
        .AddLineSegment -1259495, -25346
        .AddLineSegment -1259495, -17824
        .AddLineSegment -1260495, -17824
        .AddLineSegment -1260495, -25346
        .CloseSubPath

        .StartSubPath
        .AddLineSegment -1270495, -25346
        .AddLineSegment -1269495, -25346
        .AddLineSegment -1269495, -17824
        .AddLineSegment -1270495, -17824
        .AddLineSegment -1270495, -25346
        .CloseSubPath

        .StartSubPath
        .AddLineSegment -1280495, -25346
        .AddLineSegment -1279495, -25346
        .AddLineSegment -1279495, -17824
        .AddLineSegment -1280495, -17824
        .AddLineSegment -1280495, -25346
        .CloseSubPath

        .StartSubPath
        .AddLineSegment -1300495, -25346
        .AddLineSegment -1299495, -25346
        .AddLineSegment -1299495, -17824
        .AddLineSegment -1300495, -17824
        .AddLineSegment -1300495, -25346
        .CloseSubPath

        .StartSubPath
        .AddLineSegment -1330495, -25346
        .AddLineSegment -1329495, -25346
        .AddLineSegment -1329495, -17824
        .AddLineSegment -1330495, -17824
        .AddLineSegment -1330495, -25346
        .CloseSubPath

        .StartSubPath
        .AddLineSegment -1310495, -25346
        .AddLineSegment -1309495, -25346
        .AddLineSegment -1309495, -17824
        .AddLineSegment -1310495, -17824
        .AddLineSegment -1310495, -25346
        .CloseSubPath

        .StartSubPath
        .AddLineSegment -1320495, -25346
        .AddLineSegment -1319495, -25346
        .AddLineSegment -1319495, -17824
        .AddLineSegment -1320495, -17824
        .AddLineSegment -1320495, -25346
        .CloseSubPath

        .StartSubPath
        .AddLineSegment -1340495, -25346
        .AddLineSegment -1339495, -25346
        .AddLineSegment -1339495, -17824
        .AddLineSegment -1340495, -17824
        .AddLineSegment -1340495, -25346
        .CloseSubPath

        .StartSubPath
        .AddLineSegment -1350495, -25346
        .AddLineSegment -1349495, -25346
        .AddLineSegment -1349495, -17824
        .AddLineSegment -1350495, -17824
        .AddLineSegment -1350495, -25346
        .CloseSubPath

        .StartSubPath
        .AddLineSegment -1360495, -25346
        .AddLineSegment -1359495, -25346
        .AddLineSegment -1359495, -17824
        .AddLineSegment -1360495, -17824
        .AddLineSegment -1360495, -25346
        .CloseSubPath

        .StartSubPath
        .AddLineSegment -1380495, -25346
        .AddLineSegment -1379495, -25346
        .AddLineSegment -1379495, -17824
        .AddLineSegment -1380495, -17824
        .AddLineSegment -1380495, -25346
        .CloseSubPath

        .StartSubPath
        .AddLineSegment -1295495, -25346
        .AddLineSegment -1294495, -25346
        .AddLineSegment -1294495, -21737
        .AddLineSegment -1295495, -21737
        .AddLineSegment -1295495, -25346
        .CloseSubPath

        .StartSubPath
        .AddLineSegment -1255495, -25346
        .AddLineSegment -1254495, -25346
        .AddLineSegment -1254495, -21737
        .AddLineSegment -1255495, -21737
        .AddLineSegment -1255495, -25346
        .CloseSubPath

        .StartSubPath
        .AddLineSegment -1235495, -25346
        .AddLineSegment -1234495, -25346
        .AddLineSegment -1234495, -21737
        .AddLineSegment -1235495, -21737
        .AddLineSegment -1235495, -25346
        .CloseSubPath

        .StartSubPath
        .AddLineSegment -1245495, -25346
        .AddLineSegment -1244495, -25346
        .AddLineSegment -1244495, -21737
        .AddLineSegment -1245495, -21737
        .AddLineSegment -1245495, -25346
        .CloseSubPath

        .StartSubPath
        .AddLineSegment -1265495, -25346
        .AddLineSegment -1264495, -25346
        .AddLineSegment -1264495, -21737
        .AddLineSegment -1265495, -21737
        .AddLineSegment -1265495, -25346
        .CloseSubPath

        .StartSubPath
        .AddLineSegment -1275495, -25346
        .AddLineSegment -1274495, -25346
        .AddLineSegment -1274495, -21737
        .AddLineSegment -1275495, -21737
        .AddLineSegment -1275495, -25346
        .CloseSubPath

        .StartSubPath
        .AddLineSegment -1285495, -25346
        .AddLineSegment -1284495, -25346
        .AddLineSegment -1284495, -21737
        .AddLineSegment -1285495, -21737
        .AddLineSegment -1285495, -25346
        .CloseSubPath

        .StartSubPath
        .AddLineSegment -1305495, -25346
        .AddLineSegment -1304495, -25346
        .AddLineSegment -1304495, -21737
        .AddLineSegment -1305495, -21737
        .AddLineSegment -1305495, -25346
        .CloseSubPath

        .StartSubPath
        .AddLineSegment -1335495, -25346
        .AddLineSegment -1334495, -25346
        .AddLineSegment -1334495, -21737
        .AddLineSegment -1335495, -21737
        .AddLineSegment -1335495, -25346
        .CloseSubPath

        .StartSubPath
        .AddLineSegment -1315495, -25346
        .AddLineSegment -1314495, -25346
        .AddLineSegment -1314495, -21737
        .AddLineSegment -1315495, -21737
        .AddLineSegment -1315495, -25346
        .CloseSubPath

        .StartSubPath
        .AddLineSegment -1325495, -25346
        .AddLineSegment -1324495, -25346
        .AddLineSegment -1324495, -21737
        .AddLineSegment -1325495, -21737
        .AddLineSegment -1325495, -25346
        .CloseSubPath

        .StartSubPath
        .AddLineSegment -1345495, -25346
        .AddLineSegment -1344495, -25346
        .AddLineSegment -1344495, -21737
        .AddLineSegment -1345495, -21737
        .AddLineSegment -1345495, -25346
        .CloseSubPath

        .StartSubPath
        .AddLineSegment -1355495, -25346
        .AddLineSegment -1354495, -25346
        .AddLineSegment -1354495, -21737
        .AddLineSegment -1355495, -21737
        .AddLineSegment -1355495, -25346
        .CloseSubPath

        .StartSubPath
        .AddLineSegment -1365495, -25346
        .AddLineSegment -1364495, -25346
        .AddLineSegment -1364495, -21737
        .AddLineSegment -1365495, -21737
        .AddLineSegment -1365495, -25346
        .CloseSubPath

        .StartSubPath
        .AddLineSegment -1385495, -25346
        .AddLineSegment -1384495, -25346
        .AddLineSegment -1384495, -21737
        .AddLineSegment -1385495, -21737
        .AddLineSegment -1385495, -25346
        .CloseSubPath

        .StartSubPath
        .AddLineSegment -1395495, -25346
        .AddLineSegment -1394495, -25346
        .AddLineSegment -1394495, -21737
        .AddLineSegment -1395495, -21737
        .AddLineSegment -1395495, -25346
        .CloseSubPath

        .StartSubPath
        .AddLineSegment -1390495, -25346
        .AddLineSegment -1389495, -25346
        .AddLineSegment -1389495, -17824
        .AddLineSegment -1390495, -17824
        .AddLineSegment -1390495, -25346
        .CloseSubPath

        .StartSubPath
        .AddLineSegment -670497, -25346
        .AddLineSegment -669497, -25346
        .AddLineSegment -669497, -17824
        .AddLineSegment -670497, -17824
        .AddLineSegment -670497, -25346
        .CloseSubPath

        .StartSubPath
        .AddLineSegment -1465494, -25346
        .AddLineSegment -1464494, -25346
        .AddLineSegment -1464494, -21737
        .AddLineSegment -1465494, -21737
        .AddLineSegment -1465494, -25346
        .CloseSubPath

        .StartSubPath
        .AddLineSegment -1460494, -25346
        .AddLineSegment -1459494, -25346
        .AddLineSegment -1459494, -17824
        .AddLineSegment -1460494, -17824
        .AddLineSegment -1460494, -25346
        .CloseSubPath

        .StartSubPath
        .AddLineSegment -1420495, -25346
        .AddLineSegment -1419495, -25346
        .AddLineSegment -1419495, -17824
        .AddLineSegment -1420495, -17824
        .AddLineSegment -1420495, -25346
        .CloseSubPath

        .StartSubPath
        .AddLineSegment -1410495, -25346
        .AddLineSegment -1409495, -25346
        .AddLineSegment -1409495, -17824
        .AddLineSegment -1410495, -17824
        .AddLineSegment -1410495, -25346
        .CloseSubPath

        .StartSubPath
        .AddLineSegment -1430494, -25346
        .AddLineSegment -1429494, -25346
        .AddLineSegment -1429494, -17824
        .AddLineSegment -1430494, -17824
        .AddLineSegment -1430494, -25346
        .CloseSubPath

        .StartSubPath
        .AddLineSegment -1440495, -25346
        .AddLineSegment -1439495, -25346
        .AddLineSegment -1439495, -17824
        .AddLineSegment -1440495, -17824
        .AddLineSegment -1440495, -25346
        .CloseSubPath

        .StartSubPath
        .AddLineSegment -1450494, -25346
        .AddLineSegment -1449494, -25346
        .AddLineSegment -1449494, -17824
        .AddLineSegment -1450494, -17824
        .AddLineSegment -1450494, -25346
        .CloseSubPath

        .StartSubPath
        .AddLineSegment -1470494, -25346
        .AddLineSegment -1469494, -25346
        .AddLineSegment -1469494, -17824
        .AddLineSegment -1470494, -17824
        .AddLineSegment -1470494, -25346
        .CloseSubPath

        .StartSubPath
        .AddLineSegment -1425495, -25346
        .AddLineSegment -1424495, -25346
        .AddLineSegment -1424495, -21737
        .AddLineSegment -1425495, -21737
        .AddLineSegment -1425495, -25346
        .CloseSubPath

        .StartSubPath
        .AddLineSegment -1405495, -25346
        .AddLineSegment -1404495, -25346
        .AddLineSegment -1404495, -21737
        .AddLineSegment -1405495, -21737
        .AddLineSegment -1405495, -25346
        .CloseSubPath

        .StartSubPath
        .AddLineSegment -1415495, -25346
        .AddLineSegment -1414495, -25346
        .AddLineSegment -1414495, -21737
        .AddLineSegment -1415495, -21737
        .AddLineSegment -1415495, -25346
        .CloseSubPath

        .StartSubPath
        .AddLineSegment -1435494, -25346
        .AddLineSegment -1434494, -25346
        .AddLineSegment -1434494, -21737
        .AddLineSegment -1435494, -21737
        .AddLineSegment -1435494, -25346
        .CloseSubPath

        .StartSubPath
        .AddLineSegment -1445495, -25346
        .AddLineSegment -1444495, -25346
        .AddLineSegment -1444495, -21737
        .AddLineSegment -1445495, -21737
        .AddLineSegment -1445495, -25346
        .CloseSubPath

        .StartSubPath
        .AddLineSegment -1455494, -25346
        .AddLineSegment -1454494, -25346
        .AddLineSegment -1454494, -21737
        .AddLineSegment -1455494, -21737
        .AddLineSegment -1455494, -25346
        .CloseSubPath

        .StartSubPath
        .AddLineSegment -1475494, -25346
        .AddLineSegment -1474494, -25346
        .AddLineSegment -1474494, -21737
        .AddLineSegment -1475494, -21737
        .AddLineSegment -1475494, -25346
        .CloseSubPath

        .StartSubPath
        .AddLineSegment -1625494, -25346
        .AddLineSegment -1624494, -25346
        .AddLineSegment -1624494, -21737
        .AddLineSegment -1625494, -21737
        .AddLineSegment -1625494, -25346
        .CloseSubPath

        .StartSubPath
        .AddLineSegment -1620494, -25346
        .AddLineSegment -1619494, -25346
        .AddLineSegment -1619494, -17824
        .AddLineSegment -1620494, -17824
        .AddLineSegment -1620494, -25346
        .CloseSubPath

        .StartSubPath
        .AddLineSegment -1540494, -25346
        .AddLineSegment -1539494, -25346
        .AddLineSegment -1539494, -17824
        .AddLineSegment -1540494, -17824
        .AddLineSegment -1540494, -25346
        .CloseSubPath

        .StartSubPath
        .AddLineSegment -1500494, -25346
        .AddLineSegment -1499494, -25346
        .AddLineSegment -1499494, -17824
        .AddLineSegment -1500494, -17824
        .AddLineSegment -1500494, -25346
        .CloseSubPath

        .StartSubPath
        .AddLineSegment -1480494, -25346
        .AddLineSegment -1479494, -25346
        .AddLineSegment -1479494, -17824
        .AddLineSegment -1480494, -17824
        .AddLineSegment -1480494, -25346
        .CloseSubPath

        .StartSubPath
        .AddLineSegment -1490494, -25346
        .AddLineSegment -1489494, -25346
        .AddLineSegment -1489494, -17824
        .AddLineSegment -1490494, -17824
        .AddLineSegment -1490494, -25346
        .CloseSubPath

        .StartSubPath
        .AddLineSegment -1510494, -25346
        .AddLineSegment -1509494, -25346
        .AddLineSegment -1509494, -17824
        .AddLineSegment -1510494, -17824
        .AddLineSegment -1510494, -25346
        .CloseSubPath

        .StartSubPath
        .AddLineSegment -1520494, -25346
        .AddLineSegment -1519494, -25346
        .AddLineSegment -1519494, -17824
        .AddLineSegment -1520494, -17824
        .AddLineSegment -1520494, -25346
        .CloseSubPath

        .StartSubPath
        .AddLineSegment -1530494, -25346
        .AddLineSegment -1529494, -25346
        .AddLineSegment -1529494, -17824
        .AddLineSegment -1530494, -17824
        .AddLineSegment -1530494, -25346
        .CloseSubPath

        .StartSubPath
        .AddLineSegment -1550494, -25346
        .AddLineSegment -1549494, -25346
        .AddLineSegment -1549494, -17824
        .AddLineSegment -1550494, -17824
        .AddLineSegment -1550494, -25346
        .CloseSubPath

        .StartSubPath
        .AddLineSegment -1580494, -25346
        .AddLineSegment -1579494, -25346
        .AddLineSegment -1579494, -17824
        .AddLineSegment -1580494, -17824
        .AddLineSegment -1580494, -25346
        .CloseSubPath

        .StartSubPath
        .AddLineSegment -1560494, -25346
        .AddLineSegment -1559494, -25346
        .AddLineSegment -1559494, -17824
        .AddLineSegment -1560494, -17824
        .AddLineSegment -1560494, -25346
        .CloseSubPath

        .StartSubPath
        .AddLineSegment -1570494, -25346
        .AddLineSegment -1569494, -25346
        .AddLineSegment -1569494, -17824
        .AddLineSegment -1570494, -17824
        .AddLineSegment -1570494, -25346
        .CloseSubPath

        .StartSubPath
        .AddLineSegment -1590494, -25346
        .AddLineSegment -1589494, -25346
        .AddLineSegment -1589494, -17824
        .AddLineSegment -1590494, -17824
        .AddLineSegment -1590494, -25346
        .CloseSubPath

        .StartSubPath
        .AddLineSegment -1600494, -25346
        .AddLineSegment -1599494, -25346
        .AddLineSegment -1599494, -17824
        .AddLineSegment -1600494, -17824
        .AddLineSegment -1600494, -25346
        .CloseSubPath

        .StartSubPath
        .AddLineSegment -1610494, -25346
        .AddLineSegment -1609494, -25346
        .AddLineSegment -1609494, -17824
        .AddLineSegment -1610494, -17824
        .AddLineSegment -1610494, -25346
        .CloseSubPath

        .StartSubPath
        .AddLineSegment -1630494, -25346
        .AddLineSegment -1629494, -25346
        .AddLineSegment -1629494, -17824
        .AddLineSegment -1630494, -17824
        .AddLineSegment -1630494, -25346
        .CloseSubPath

        .StartSubPath
        .AddLineSegment -1545494, -25346
        .AddLineSegment -1544494, -25346
        .AddLineSegment -1544494, -21737
        .AddLineSegment -1545494, -21737
        .AddLineSegment -1545494, -25346
        .CloseSubPath

        .StartSubPath
        .AddLineSegment -1505494, -25346
        .AddLineSegment -1504494, -25346
        .AddLineSegment -1504494, -21737
        .AddLineSegment -1505494, -21737
        .AddLineSegment -1505494, -25346
        .CloseSubPath

        .StartSubPath
        .AddLineSegment -1485494, -25346
        .AddLineSegment -1484494, -25346
        .AddLineSegment -1484494, -21737
        .AddLineSegment -1485494, -21737
        .AddLineSegment -1485494, -25346
        .CloseSubPath

        .StartSubPath
        .AddLineSegment -1495494, -25346
        .AddLineSegment -1494494, -25346
        .AddLineSegment -1494494, -21737
        .AddLineSegment -1495494, -21737
        .AddLineSegment -1495494, -25346
        .CloseSubPath

        .StartSubPath
        .AddLineSegment -1515494, -25346
        .AddLineSegment -1514494, -25346
        .AddLineSegment -1514494, -21737
        .AddLineSegment -1515494, -21737
        .AddLineSegment -1515494, -25346
        .CloseSubPath

        .StartSubPath
        .AddLineSegment -1525494, -25346
        .AddLineSegment -1524494, -25346
        .AddLineSegment -1524494, -21737
        .AddLineSegment -1525494, -21737
        .AddLineSegment -1525494, -25346
        .CloseSubPath

        .StartSubPath
        .AddLineSegment -1535494, -25346
        .AddLineSegment -1534494, -25346
        .AddLineSegment -1534494, -21737
        .AddLineSegment -1535494, -21737
        .AddLineSegment -1535494, -25346
        .CloseSubPath

        .StartSubPath
        .AddLineSegment -1555494, -25346
        .AddLineSegment -1554494, -25346
        .AddLineSegment -1554494, -21737
        .AddLineSegment -1555494, -21737
        .AddLineSegment -1555494, -25346
        .CloseSubPath

        .StartSubPath
        .AddLineSegment -1585494, -25346
        .AddLineSegment -1584494, -25346
        .AddLineSegment -1584494, -21737
        .AddLineSegment -1585494, -21737
        .AddLineSegment -1585494, -25346
        .CloseSubPath

        .StartSubPath
        .AddLineSegment -1565494, -25346
        .AddLineSegment -1564494, -25346
        .AddLineSegment -1564494, -21737
        .AddLineSegment -1565494, -21737
        .AddLineSegment -1565494, -25346
        .CloseSubPath

        .StartSubPath
        .AddLineSegment -1575494, -25346
        .AddLineSegment -1574494, -25346
        .AddLineSegment -1574494, -21737
        .AddLineSegment -1575494, -21737
        .AddLineSegment -1575494, -25346
        .CloseSubPath

        .StartSubPath
        .AddLineSegment -1595494, -25346
        .AddLineSegment -1594494, -25346
        .AddLineSegment -1594494, -21737
        .AddLineSegment -1595494, -21737
        .AddLineSegment -1595494, -25346
        .CloseSubPath

        .StartSubPath
        .AddLineSegment -1605494, -25346
        .AddLineSegment -1604494, -25346
        .AddLineSegment -1604494, -21737
        .AddLineSegment -1605494, -21737
        .AddLineSegment -1605494, -25346
        .CloseSubPath

        .StartSubPath
        .AddLineSegment -1615494, -25346
        .AddLineSegment -1614494, -25346
        .AddLineSegment -1614494, -21737
        .AddLineSegment -1615494, -21737
        .AddLineSegment -1615494, -25346
        .CloseSubPath

        .StartSubPath
        .AddLineSegment -1635494, -25346
        .AddLineSegment -1634494, -25346
        .AddLineSegment -1634494, -21737
        .AddLineSegment -1635494, -21737
        .AddLineSegment -1635494, -25346
        .CloseSubPath

        .StartSubPath
        .AddLineSegment -1660494, -25346
        .AddLineSegment -1659494, -25346
        .AddLineSegment -1659494, -17824
        .AddLineSegment -1660494, -17824
        .AddLineSegment -1660494, -25346
        .CloseSubPath

        .StartSubPath
        .AddLineSegment -1640494, -25346
        .AddLineSegment -1639494, -25346
        .AddLineSegment -1639494, -17824
        .AddLineSegment -1640494, -17824
        .AddLineSegment -1640494, -25346
        .CloseSubPath

        .StartSubPath
        .AddLineSegment -1650494, -25346
        .AddLineSegment -1649494, -25346
        .AddLineSegment -1649494, -17824
        .AddLineSegment -1650494, -17824
        .AddLineSegment -1650494, -25346
        .CloseSubPath

        .StartSubPath
        .AddLineSegment -1665494, -25346
        .AddLineSegment -1664494, -25346
        .AddLineSegment -1664494, -21737
        .AddLineSegment -1665494, -21737
        .AddLineSegment -1665494, -25346
        .CloseSubPath

        .StartSubPath
        .AddLineSegment -1645494, -25346
        .AddLineSegment -1644494, -25346
        .AddLineSegment -1644494, -21737
        .AddLineSegment -1645494, -21737
        .AddLineSegment -1645494, -25346
        .CloseSubPath

        .StartSubPath
        .AddLineSegment -1655494, -25346
        .AddLineSegment -1654494, -25346
        .AddLineSegment -1654494, -21737
        .AddLineSegment -1655494, -21737
        .AddLineSegment -1655494, -25346
        .CloseSubPath

        .StartSubPath
        .AddLineSegment -1400495, -25346
        .AddLineSegment -1399495, -25346
        .AddLineSegment -1399495, -17824
        .AddLineSegment -1400495, -17824
        .AddLineSegment -1400495, -25346
        .CloseSubPath

        Set .brush = New waBrush
        .brush.SetColor RGB(114, 114, 113), 100
    End With
    MyCurves.Add path
    InitStandardPlotter1600_2Dmodel_2 MyCurves

    StandardPlotter1600_2DModel_Initialized = True
End Sub

' -----------------------------------------------------------------------------
' InitStandardPlotter1600_2Dmodel_2 (Sub), source lines 7111-8381
' -----------------------------------------------------------------------------
Private Sub InitStandardPlotter1600_2Dmodel_2(MyCurves As Collection)
    Dim path As waPath
    Set path = New waPath
    With path
        .StartSubPath
        .AddLineSegment -1670499, 119627
        .AddLineSegment 70499, 119627
        .AddLineSegment 70499, 91572
        .AddLineSegment -1670499, 91572
        .AddLineSegment -1670499, 119627
        .CloseSubPath

        Set .brush = New waBrush
        .brush.SetLinearGradient -800000, 92000, -800000, 120000, RGB(254, 254, 254), RGB(233, 233, 233), WrapModeTileFlipXY
    End With
    MyCurves.Add path

    Set path = New waPath
    With path
        .StartSubPath
        .AddLineSegment -1670499, 78872
        .AddLineSegment 70499, 78872
        .AddLineSegment 70499, 91837
        .AddLineSegment -1670499, 91837
        .AddLineSegment -1670499, 78872
        .CloseSubPath

        Set .brush = New waBrush
        .brush.SetColor RGB(254, 254, 254), 100
    End With
    MyCurves.Add path

    Set path = New waPath
    With path
        .StartSubPath
        .AddLineSegment -1652398, 81139
        .AddBezierSegment -1652398, 81139, -1650177, 81139, -1648376, 82939, -1648376, 85160
        .AddBezierSegment -1648376, 85160, -1648376, 87381, -1650177, 89182, -1652398, 89182
        .AddBezierSegment -1652398, 89182, -1654619, 89182, -1656420, 87381, -1656420, 85160
        .AddBezierSegment -1656420, 85160, -1656420, 82939, -1654619, 81139, -1652398, 81139
        .CloseSubPath

        Set .brush = New waBrush
        .brush.SetLinearGradient -1652000, 85000, -1648000, 85000, RGB(217, 218, 218), RGB(242, 239, 238), WrapModeTileFlipXY
    End With
    MyCurves.Add path

    Set path = New waPath
    With path
        .StartSubPath
        .AddLineSegment -1653328, 86273
        .AddLineSegment -1652897, 86522
        .AddLineSegment -1652307, 85500
        .AddLineSegment -1651285, 86090
        .AddLineSegment -1651036, 85659
        .AddLineSegment -1652059, 85069
        .AddLineSegment -1651468, 84047
        .AddLineSegment -1651899, 83799
        .AddLineSegment -1652489, 84821
        .AddLineSegment -1653511, 84231
        .AddLineSegment -1653760, 84661
        .AddLineSegment -1652738, 85251
        .AddLineSegment -1653328, 86273
        .CloseSubPath

        Set .brush = New waBrush
        .brush.SetColor RGB(197, 198, 198), 100
    End With
    MyCurves.Add path

    Set path = New waPath
    With path
        .StartSubPath
        .AddLineSegment 48450, 82087
        .AddBezierSegment 48450, 82087, 50147, 80654, 52684, 80869, 54117, 82566
        .AddBezierSegment 54117, 82566, 55550, 84263, 55335, 86801, 53638, 88234
        .AddBezierSegment 53638, 88234, 51940, 89666, 49403, 89452, 47970, 87754
        .AddBezierSegment 47970, 87754, 46538, 86057, 46752, 83520, 48450, 82087
        .CloseSubPath

        Set .brush = New waBrush
        .brush.SetLinearGradient 51000, 85000, 54000, 83000, RGB(217, 218, 218), RGB(242, 239, 238), WrapModeTileFlipXY
    End With
    MyCurves.Add path

    Set path = New waPath
    With path
        .StartSubPath
        .AddLineSegment 51051, 86611
        .AddLineSegment 51540, 86523
        .AddLineSegment 51332, 85361
        .AddLineSegment 52494, 85153
        .AddLineSegment 52406, 84663
        .AddLineSegment 51244, 84872
        .AddLineSegment 51036, 83710
        .AddLineSegment 50547, 83798
        .AddLineSegment 50755, 84960
        .AddLineSegment 49593, 85168
        .AddLineSegment 49681, 85657
        .AddLineSegment 50843, 85449
        .AddLineSegment 51051, 86611
        .CloseSubPath

        Set .brush = New waBrush
        .brush.SetColor RGB(197, 198, 198), 100
    End With
    MyCurves.Add path

    Set path = New waPath
    With path
        .StartSubPath
        .AddLineSegment -1670499, -17824
        .AddLineSegment 70500, -17824
        .AddLineSegment 70500, 0
        .AddLineSegment -1670499, 0
        .AddLineSegment -1670499, -17824
        .CloseSubPath

        Set .brush = New waBrush
        .brush.SetLinearGradient -800000, -18000, -800000, 0, RGB(164, 185, 195), RGB(148, 148, 148), WrapModeTileFlipXY
    End With
    MyCurves.Add path

    Set path = New waPath
    With path
        .StartSubPath
        .AddLineSegment -1629014, 27225
        .AddLineSegment 70499, 27225
        .AddLineSegment 70499, 22145
        .AddLineSegment -1629014, 22145
        .AddLineSegment -1629014, 27225
        .CloseSubPath

        Set .brush = New waBrush
        .brush.SetLinearGradient -779000, 25000, -779000, 20000, RGB(43, 42, 41), RGB(194, 194, 194), WrapModeTileFlipXY
    End With
    MyCurves.Add path

    Set path = New waPath
    With path
        .StartSubPath
        .AddLineSegment 70499, 66171
        .AddLineSegment -1670499, 66171
        .AddLineSegment -1670499, 27225
        .AddLineSegment 70499, 27225
        .AddLineSegment 70499, 66171
        .CloseSubPath

        Set .brush = New waBrush
        .brush.SetColor RGB(67, 66, 66), 100
    End With
    MyCurves.Add path

    Set path = New waPath
    With path
        .StartSubPath
        .AddLineSegment -1670499, 46250
        .AddLineSegment 70500, 46250
        .AddLineSegment 70500, 36538
        .AddLineSegment -1670499, 36538
        .AddLineSegment -1670499, 46250
        .CloseSubPath

        Set .brush = New waBrush
        .brush.SetColor RGB(43, 42, 41), 100
    End With
    MyCurves.Add path

    Set path = New waPath
    With path
        .StartSubPath
        .AddLineSegment -627200, 32622
        .AddLineSegment -579786, 32622
        .AddLineSegment -579786, 27225
        .AddLineSegment -627200, 27225
        .AddLineSegment -627200, 32622
        .CloseSubPath

        .StartSubPath
        .AddLineSegment -1627212, 32622
        .AddLineSegment -1579798, 32622
        .AddLineSegment -1579798, 27225
        .AddLineSegment -1627212, 27225
        .AddLineSegment -1627212, 32622
        .CloseSubPath

        .StartSubPath
        .AddLineSegment -816052, 32622
        .AddLineSegment -768638, 32622
        .AddLineSegment -768638, 27225
        .AddLineSegment -816052, 27225
        .AddLineSegment -816052, 32622
        .CloseSubPath

        .StartSubPath
        .AddLineSegment -1010468, 32622
        .AddLineSegment -963054, 32622
        .AddLineSegment -963054, 27225
        .AddLineSegment -1010468, 27225
        .AddLineSegment -1010468, 32622
        .CloseSubPath

        .StartSubPath
        .AddLineSegment -1189961, 32622
        .AddLineSegment -1142548, 32622
        .AddLineSegment -1142548, 27225
        .AddLineSegment -1189961, 27225
        .AddLineSegment -1189961, 32622
        .CloseSubPath

        .StartSubPath
        .AddLineSegment -1382684, 32622
        .AddLineSegment -1335271, 32622
        .AddLineSegment -1335271, 27225
        .AddLineSegment -1382684, 27225
        .AddLineSegment -1382684, 32622
        .CloseSubPath

        .StartSubPath
        .AddLineSegment -8233, 32622
        .AddLineSegment 39180, 32622
        .AddLineSegment 39180, 27225
        .AddLineSegment -8233, 27225
        .AddLineSegment -8233, 32622
        .CloseSubPath

        .StartSubPath
        .AddLineSegment -112055, 32622
        .AddLineSegment -64642, 32622
        .AddLineSegment -64642, 27225
        .AddLineSegment -112055, 27225
        .AddLineSegment -112055, 32622
        .CloseSubPath

        .StartSubPath
        .AddLineSegment -189949, 32622
        .AddLineSegment -142536, 32622
        .AddLineSegment -142536, 27225
        .AddLineSegment -189949, 27225
        .AddLineSegment -189949, 32622
        .CloseSubPath

        .StartSubPath
        .AddLineSegment -382672, 32622
        .AddLineSegment -335258, 32622
        .AddLineSegment -335258, 27225
        .AddLineSegment -382672, 27225
        .AddLineSegment -382672, 32622
        .CloseSubPath

        Set .brush = New waBrush
        .brush.SetColor RGB(254, 254, 254), 100
    End With
    MyCurves.Add path

    Set path = New waPath
    With path
        .StartSubPath
        .AddLineSegment -814076, 0
        .AddLineSegment -785924, 0
        .AddLineSegment -785924, 6430
        .AddLineSegment -789628, 6430
        .AddLineSegment -789628, 22145
        .AddLineSegment -810372, 22145
        .AddLineSegment -810372, 6430
        .AddLineSegment -814076, 6430
        .AddLineSegment -814076, 0
        .CloseSubPath

        Set .brush = New waBrush
        .brush.SetColor RGB(43, 42, 41), 100
    End With
    MyCurves.Add path

    Set path = New waPath
    With path
        .StartSubPath
        .AddLineSegment -1625000, 22145
        .AddLineSegment -1599970, 22145
        .AddLineSegment -1599970, 2567
        .AddBezierSegment -1599970, 2567, -1599970, 2014, -1600423, 1562, -1600976, 1562
        .AddLineSegment -1605738, 1562
        .AddBezierSegment -1605738, 1562, -1606291, 1562, -1606744, 2014, -1606744, 2567
        .AddLineSegment -1606744, 9711
        .AddLineSegment -1618227, 9711
        .AddLineSegment -1618227, 2567
        .AddBezierSegment -1618227, 2567, -1618227, 2014, -1618679, 1562, -1619232, 1562
        .AddLineSegment -1623995, 1562
        .AddBezierSegment -1623995, 1562, -1624547, 1562, -1625000, 2014, -1625000, 2567
        .AddBezierSegment -1625000, 2567, -1625000, 9093, -1625000, 15619, -1625000, 22145
        .CloseSubPath

        Set .brush = New waBrush
        .brush.SetColor RGB(43, 42, 41), 100
    End With
    MyCurves.Add path

    Set path = New waPath
    With path
        .StartSubPath
        .AddLineSegment -1615528, 0
        .AddLineSegment -1609389, 0
        .AddBezierSegment -1609389, 0, -1608109, 0, -1607061, 1048, -1607061, 2328
        .AddLineSegment -1607061, 10716
        .AddBezierSegment -1607061, 10716, -1607061, 11997, -1608109, 13045, -1609389, 13045
        .AddLineSegment -1615528, 13045
        .AddBezierSegment -1615528, 13045, -1616808, 13045, -1617856, 11997, -1617856, 10716
        .AddLineSegment -1617856, 2328
        .AddBezierSegment -1617856, 2328, -1617856, 1048, -1616808, 0, -1615528, 0
        .CloseSubPath

        Set .brush = New waBrush
        .brush.SetLinearGradient -1612000, -7000, -1612000, 7000, RGB(43, 42, 41), RGB(167, 165, 165), WrapModeTileFlipXY
    End With
    MyCurves.Add path

    Set path = New waPath
    With path
        .StartSubPath
        .AddLineSegment -41, 22145
        .AddLineSegment 24988, 22145
        .AddLineSegment 24988, 2567
        .AddBezierSegment 24988, 2567, 24988, 2014, 24536, 1562, 23983, 1562
        .AddLineSegment 19220, 1562
        .AddBezierSegment 19220, 1562, 18667, 1562, 18215, 2014, 18215, 2567
        .AddLineSegment 18215, 9711
        .AddLineSegment 6732, 9711
        .AddLineSegment 6732, 2567
        .AddBezierSegment 6732, 2567, 6732, 2014, 6279, 1562, 5726, 1562
        .AddLineSegment 964, 1562
        .AddBezierSegment 964, 1562, 411, 1562, -41, 2014, -41, 2567
        .AddBezierSegment -41, 2567, -41, 9093, -41, 15619, -41, 22145
        .CloseSubPath

        Set .brush = New waBrush
        .brush.SetColor RGB(43, 42, 41), 100
    End With
    MyCurves.Add path

    Set path = New waPath
    With path
        .StartSubPath
        .AddLineSegment 9431, 0
        .AddLineSegment 15569, 0
        .AddBezierSegment 15569, 0, 16850, 0, 17898, 1048, 17898, 2328
        .AddLineSegment 17898, 10716
        .AddBezierSegment 17898, 10716, 17898, 11997, 16850, 13045, 15569, 13045
        .AddLineSegment 9431, 13045
        .AddBezierSegment 9431, 13045, 8150, 13045, 7103, 11997, 7103, 10716
        .AddLineSegment 7103, 2328
        .AddBezierSegment 7103, 2328, 7103, 1048, 8150, 0, 9431, 0
        .CloseSubPath

        Set .brush = New waBrush
        .brush.SetLinearGradient 12000, -7000, 12000, 7000, RGB(43, 42, 41), RGB(167, 165, 165), WrapModeTileFlipXY
    End With
    MyCurves.Add path

    Set path = New waPath
    With path
        .StartSubPath
        .AddLineSegment -165061, 118871
        .AddLineSegment 20570, 118871
        .AddBezierSegment 20570, 118871, 29934, 118871, 37547, 116207, 37547, 112929
        .AddLineSegment 37547, 97513
        .AddBezierSegment 37547, 97513, 37547, 94236, 29934, 91572, 20570, 91572
        .AddLineSegment -206760, 91572
        .AddBezierSegment -206760, 91572, -216742, 91572, -225162, 94746, -223564, 98279
        .AddBezierSegment -223564, 98279, -220684, 104649, -213054, 110667, -198232, 114700
        .AddBezierSegment -198232, 114700, -188992, 117214, -176851, 118871, -165061, 118871
        .CloseSubPath

        Set .brush = New waBrush
        .brush.SetColor RGB(235, 236, 236), 100
    End With
    MyCurves.Add path

    Set path = New waPath
    With path
        .StartSubPath
        .AddLineSegment -152785, 113971
        .AddLineSegment 9564, 113971
        .AddBezierSegment 9564, 113971, 11201, 113971, 12540, 113502, 12540, 112929
        .AddLineSegment 12540, 97513
        .AddBezierSegment 12540, 97513, 12540, 96941, 11201, 96472, 9564, 96472
        .AddLineSegment -194483, 96472
        .AddBezierSegment -194483, 96472, -196120, 96472, -197680, 97026, -197460, 97513
        .AddBezierSegment -197460, 97513, -191321, 111085, -164899, 113971, -152785, 113971
        .CloseSubPath

        Set .brush = New waBrush
        .brush.SetColor RGB(43, 42, 41), 100
    End With
    MyCurves.Add path

    Set path = New waPath
    With path
        .StartSubPath
        .AddLineSegment -22005, 99119
        .AddLineSegment -11407, 99119
        .AddBezierSegment -11407, 99119, -10978, 99119, -10628, 99242, -10628, 99392
        .AddLineSegment -10628, 102429
        .AddLineSegment -16706, 104369
        .AddLineSegment -22785, 102429
        .AddLineSegment -22785, 99392
        .AddBezierSegment -22785, 99392, -22785, 99242, -22434, 99119, -22005, 99119
        .CloseSubPath

        Set .brush = New waBrush
        .brush.SetColor RGB(162, 217, 247), 100
    End With
    MyCurves.Add path

    Set path = New waPath
    With path
        .StartSubPath
        .AddLineSegment -16606, 100713
        .AddLineSegment -19755, 102201
        .AddLineSegment -13458, 102201
        .AddLineSegment -16606, 100713
        .CloseSubPath

        Set .brush = New waBrush
        .brush.SetColor RGB(43, 42, 41), 100
    End With
    MyCurves.Add path

    Set path = New waPath
    With path
        .StartSubPath
        .AddLineSegment -34115, 107067
        .AddLineSegment -34115, 103358
        .AddBezierSegment -34115, 103358, -34115, 103207, -33765, 103085, -33336, 103085
        .AddLineSegment -24657, 103085
        .AddLineSegment -19116, 105212
        .AddLineSegment -24657, 107340
        .AddLineSegment -33336, 107340
        .AddBezierSegment -33336, 107340, -33765, 107340, -34115, 107217, -34115, 107067
        .CloseSubPath

        Set .brush = New waBrush
        .brush.SetColor RGB(162, 217, 247), 100
    End With
    MyCurves.Add path

    Set path = New waPath
    With path
        .StartSubPath
        .AddLineSegment -29562, 105177
        .AddLineSegment -25309, 106279
        .AddLineSegment -25309, 104075
        .AddLineSegment -29562, 105177
        .CloseSubPath

        Set .brush = New waBrush
        .brush.SetColor RGB(43, 42, 41), 100
    End With
    MyCurves.Add path

    Set path = New waPath
    With path
        .StartSubPath
        .AddLineSegment -22005, 111305
        .AddLineSegment -11407, 111305
        .AddBezierSegment -11407, 111305, -10978, 111305, -10628, 111183, -10628, 111032
        .AddLineSegment -10628, 107995
        .AddLineSegment -16706, 106055
        .AddLineSegment -22785, 107995
        .AddLineSegment -22785, 111032
        .AddBezierSegment -22785, 111032, -22785, 111182, -22434, 111305, -22005, 111305
        .CloseSubPath

        Set .brush = New waBrush
        .brush.SetColor RGB(162, 217, 247), 100
    End With
    MyCurves.Add path

    Set path = New waPath
    With path
        .StartSubPath
        .AddLineSegment -16606, 109711
        .AddLineSegment -19755, 108223
        .AddLineSegment -13458, 108223
        .AddLineSegment -16606, 109711
        .CloseSubPath

        Set .brush = New waBrush
        .brush.SetColor RGB(43, 42, 41), 100
    End With
    MyCurves.Add path

    Set path = New waPath
    With path
        .StartSubPath
        .AddLineSegment 703, 107067
        .AddLineSegment 703, 103358
        .AddBezierSegment 703, 103358, 703, 103207, 353, 103085, -76, 103085
        .AddLineSegment -8755, 103085
        .AddLineSegment -14297, 105212
        .AddLineSegment -8755, 107340
        .AddLineSegment -76, 107340
        .AddBezierSegment -76, 107340, 352, 107340, 703, 107217, 703, 107067
        .CloseSubPath

        Set .brush = New waBrush
        .brush.SetColor RGB(162, 217, 247), 100
    End With
    MyCurves.Add path

    Set path = New waPath
    With path
        .StartSubPath
        .AddLineSegment -3850, 105177
        .AddLineSegment -8103, 106279
        .AddLineSegment -8103, 104075
        .AddLineSegment -3850, 105177
        .CloseSubPath

        Set .brush = New waBrush
        .brush.SetColor RGB(43, 42, 41), 100
    End With
    MyCurves.Add path

    Set path = New waPath
    With path
        .StartSubPath
        .AddLineSegment -49568, 100274
        .AddBezierSegment -49568, 100274, -45978, 100274, -43068, 101293, -43068, 102549
        .AddBezierSegment -43068, 102549, -43068, 103806, -45978, 104824, -49568, 104824
        .AddBezierSegment -49568, 104824, -53158, 104824, -56068, 103806, -56068, 102549
        .AddBezierSegment -56068, 102549, -56068, 101293, -53158, 100274, -49568, 100274
        .CloseSubPath

        Set .brush = New waBrush
        .brush.SetColor RGB(235, 236, 236), 100
    End With
    MyCurves.Add path

    Set path = New waPath
    With path
        .StartSubPath
        .AddLineSegment -69042, 100274
        .AddBezierSegment -69042, 100274, -65452, 100274, -62542, 101293, -62542, 102549
        .AddBezierSegment -62542, 102549, -62542, 103806, -65452, 104824, -69042, 104824
        .AddBezierSegment -69042, 104824, -72632, 104824, -75542, 103806, -75542, 102549
        .AddBezierSegment -75542, 102549, -75542, 101293, -72632, 100274, -69042, 100274
        .CloseSubPath

        Set .brush = New waBrush
        .brush.SetColor RGB(235, 236, 236), 100
    End With
    MyCurves.Add path

    Set path = New waPath
    With path
        .StartSubPath
        .AddLineSegment -88515, 100274
        .AddBezierSegment -88515, 100274, -84925, 100274, -82015, 101293, -82015, 102549
        .AddBezierSegment -82015, 102549, -82015, 103806, -84925, 104824, -88515, 104824
        .AddBezierSegment -88515, 104824, -92105, 104824, -95015, 103806, -95015, 102549
        .AddBezierSegment -95015, 102549, -95015, 101293, -92105, 100274, -88515, 100274
        .CloseSubPath

        Set .brush = New waBrush
        .brush.SetColor RGB(235, 236, 236), 100
    End With
    MyCurves.Add path

    Set path = New waPath
    With path
        .StartSubPath
        .AddLineSegment -108200, 100274
        .AddBezierSegment -108200, 100274, -104610, 100274, -101700, 101293, -101700, 102549
        .AddBezierSegment -101700, 102549, -101700, 103806, -104610, 104824, -108200, 104824
        .AddBezierSegment -108200, 104824, -111790, 104824, -114700, 103806, -114700, 102549
        .AddBezierSegment -114700, 102549, -114700, 101293, -111790, 100274, -108200, 100274
        .CloseSubPath

        Set .brush = New waBrush
        .brush.SetColor RGB(235, 236, 236), 100
    End With
    MyCurves.Add path

    Set path = New waPath
    With path
        .StartSubPath
        .AddLineSegment -127674, 100274
        .AddBezierSegment -127674, 100274, -124084, 100274, -121174, 101293, -121174, 102549
        .AddBezierSegment -121174, 102549, -121174, 103806, -124084, 104824, -127674, 104824
        .AddBezierSegment -127674, 104824, -131263, 104824, -134173, 103806, -134173, 102549
        .AddBezierSegment -134173, 102549, -134173, 101293, -131263, 100274, -127674, 100274
        .CloseSubPath

        Set .brush = New waBrush
        .brush.SetColor RGB(235, 236, 236), 100
    End With
    MyCurves.Add path

    Set path = New waPath
    With path
        .StartSubPath
        .AddLineSegment -146724, 100274
        .AddBezierSegment -146724, 100274, -143134, 100274, -140224, 101293, -140224, 102549
        .AddBezierSegment -140224, 102549, -140224, 103806, -143134, 104824, -146724, 104824
        .AddBezierSegment -146724, 104824, -150313, 104824, -153224, 103806, -153224, 102549
        .AddBezierSegment -153224, 102549, -153224, 101293, -150313, 100274, -146724, 100274
        .CloseSubPath

        Set .brush = New waBrush
        .brush.SetColor RGB(235, 236, 236), 100
    End With
    MyCurves.Add path

    Set path = New waPath
    With path
        .StartSubPath
        .AddLineSegment -161310, 107587
        .AddBezierSegment -161310, 107587, -160205, 107587, -159310, 107900, -159310, 108287
        .AddBezierSegment -159310, 108287, -159310, 108673, -160205, 108987, -161310, 108987
        .AddBezierSegment -161310, 108987, -162415, 108987, -163310, 108673, -163310, 108287
        .AddBezierSegment -163310, 108287, -163310, 107900, -162415, 107587, -161310, 107587
        .CloseSubPath

        Set .brush = New waBrush
        .brush.SetColor RGB(219, 226, 133), 100
    End With
    MyCurves.Add path

    Set path = New waPath
    With path
        .StartSubPath
        .AddLineSegment -147147, 107587
        .AddBezierSegment -147147, 107587, -146042, 107587, -145147, 107900, -145147, 108287
        .AddBezierSegment -145147, 108287, -145147, 108673, -146042, 108987, -147147, 108987
        .AddBezierSegment -147147, 108987, -148251, 108987, -149147, 108673, -149147, 108287
        .AddBezierSegment -149147, 108287, -149147, 107900, -148251, 107587, -147147, 107587
        .CloseSubPath

        Set .brush = New waBrush
        .brush.SetColor RGB(219, 226, 133), 100
    End With
    MyCurves.Add path

    Set path = New waPath
    With path
        .StartSubPath
        .AddLineSegment -822356, 81964
        .AddLineSegment -822356, 88026
        .AddLineSegment -820393, 88026
        .AddBezierSegment -820393, 88026, -819649, 88026, -819165, 87997, -818940, 87936
        .AddBezierSegment -818940, 87936, -818593, 87844, -818302, 87647, -818068, 87342
        .AddBezierSegment -818068, 87342, -817833, 87037, -817716, 86643, -817716, 86161
        .AddBezierSegment -817716, 86161, -817716, 85788, -817784, 85476, -817918, 85222
        .AddBezierSegment -817918, 85222, -818053, 84968, -818224, 84769, -818432, 84624
        .AddBezierSegment -818432, 84624, -818639, 84480, -818851, 84384, -819065, 84338
        .AddBezierSegment -819065, 84338, -819357, 84280, -819779, 84250, -820332, 84250
        .AddLineSegment -821128, 84250
        .AddLineSegment -821128, 81964
        .AddLineSegment -822356, 81964
        .CloseSubPath

        .StartSubPath
        .AddLineSegment -821128, 87002
        .AddLineSegment -821128, 85275
        .AddLineSegment -820461, 85275
        .AddBezierSegment -820461, 85275, -819980, 85275, -819659, 85307, -819497, 85371
        .AddBezierSegment -819497, 85371, -819335, 85434, -819208, 85533, -819116, 85670
        .AddBezierSegment -819116, 85670, -819024, 85805, -818978, 85962, -818978, 86142
        .AddBezierSegment -818978, 86142, -818978, 86364, -819043, 86546, -819171, 86690
        .AddBezierSegment -819171, 86690, -819301, 86834, -819463, 86924, -819662, 86961
        .AddBezierSegment -819662, 86961, -819808, 86988, -820100, 87002, -820540, 87002
        .AddLineSegment -821128, 87002
        .CloseSubPath

        .StartSubPath
        .AddLineSegment -816675, 81964
        .AddLineSegment -816675, 87975
        .AddLineSegment -815447, 87975
        .AddLineSegment -815447, 82989
        .AddLineSegment -812408, 82989
        .AddLineSegment -812408, 81964
        .AddLineSegment -816675, 81964
        .CloseSubPath

        .StartSubPath
        .AddLineSegment -811791, 84959
        .AddBezierSegment -811791, 84959, -811791, 85577, -811698, 86096, -811513, 86516
        .AddBezierSegment -811513, 86516, -811376, 86825, -811187, 87103, -810949, 87349
        .AddBezierSegment -810949, 87349, -810710, 87594, -810449, 87776, -810165, 87896
        .AddBezierSegment -810165, 87896, -809787, 88056, -809351, 88136, -808857, 88136
        .AddBezierSegment -808857, 88136, -807963, 88136, -807249, 87858, -806711, 87304
        .AddBezierSegment -806711, 87304, -806175, 86748, -805907, 85976, -805907, 84987
        .AddBezierSegment -805907, 84987, -805907, 84006, -806173, 83238, -806705, 82685
        .AddBezierSegment -806705, 82685, -807237, 82131, -807950, 81854, -808840, 81854
        .AddBezierSegment -808840, 81854, -809742, 81854, -810460, 82129, -810992, 82681
        .AddBezierSegment -810992, 82681, -811524, 83231, -811791, 83990, -811791, 84959
        .CloseSubPath

        .StartSubPath
        .AddLineSegment -810529, 84999
        .AddBezierSegment -810529, 84999, -810529, 84312, -810370, 83791, -810053, 83436
        .AddBezierSegment -810053, 83436, -809735, 83082, -809332, 82904, -808843, 82904
        .AddBezierSegment -808843, 82904, -808354, 82904, -807953, 83080, -807638, 83432
        .AddBezierSegment -807638, 83432, -807325, 83784, -807168, 84312, -807168, 85016
        .AddBezierSegment -807168, 85016, -807168, 85712, -807321, 86231, -807627, 86573
        .AddBezierSegment -807627, 86573, -807932, 86916, -808337, 87087, -808843, 87087
        .AddBezierSegment -808843, 87087, -809348, 87087, -809756, 86913, -810065, 86567
        .AddBezierSegment -810065, 86567, -810374, 86220, -810529, 85698, -810529, 84999
        .CloseSubPath

        .StartSubPath
        .AddLineSegment -803588, 81964
        .AddLineSegment -803588, 87002
        .AddLineSegment -805383, 87002
        .AddLineSegment -805383, 88026
        .AddLineSegment -800574, 88026
        .AddLineSegment -800574, 87002
        .AddLineSegment -802361, 87002
        .AddLineSegment -802361, 81964
        .AddLineSegment -803588, 81964
        .CloseSubPath

        .StartSubPath
        .AddLineSegment -795902, 84199
        .AddLineSegment -794716, 83835
        .AddBezierSegment -794716, 83835, -794899, 83168, -795202, 82672, -795625, 82348
        .AddBezierSegment -795625, 82348, -796049, 82025, -796585, 81862, -797235, 81862
        .AddBezierSegment -797235, 81862, -798041, 81862, -798703, 82138, -799222, 82688
        .AddBezierSegment -799222, 82688, -799740, 83238, -800000, 83990, -800000, 84946
        .AddBezierSegment -800000, 84946, -800000, 85955, -799739, 86739, -799218, 87298
        .AddBezierSegment -799218, 87298, -798697, 87857, -798011, 88136, -797162, 88136
        .AddBezierSegment -797162, 88136, -796420, 88136, -795817, 87916, -795354, 87477
        .AddBezierSegment -795354, 87477, -795078, 87218, -794872, 86845, -794734, 86358
        .AddLineSegment -795944, 86070
        .AddBezierSegment -795944, 86070, -796016, 86384, -796166, 86632, -796393, 86814
        .AddBezierSegment -796393, 86814, -796622, 86996, -796898, 87087, -797223, 87087
        .AddBezierSegment -797223, 87087, -797673, 87087, -798038, 86926, -798318, 86602
        .AddBezierSegment -798318, 86602, -798598, 86281, -798738, 85759, -798738, 85036
        .AddBezierSegment -798738, 85036, -798738, 84271, -798600, 83725, -798325, 83401
        .AddBezierSegment -798325, 83401, -798048, 83075, -797690, 82912, -797248, 82912
        .AddBezierSegment -797248, 82912, -796922, 82912, -796641, 83015, -796407, 83223
        .AddBezierSegment -796407, 83223, -796173, 83430, -796004, 83755, -795902, 84199
        .CloseSubPath

        .StartSubPath
        .AddLineSegment -788204, 81964
        .AddLineSegment -789526, 81964
        .AddLineSegment -790050, 83344
        .AddLineSegment -792481, 83344
        .AddLineSegment -792985, 81964
        .AddLineSegment -794283, 81964
        .AddLineSegment -791930, 88026
        .AddLineSegment -790633, 88026
        .AddLineSegment -788204, 81964
        .CloseSubPath

        .StartSubPath
        .AddLineSegment -790448, 84369
        .AddLineSegment -791290, 86612
        .AddLineSegment -792109, 84369
        .AddLineSegment -790448, 84369
        .CloseSubPath

        .StartSubPath
        .AddLineSegment -787517, 81964
        .AddLineSegment -787517, 87975
        .AddLineSegment -786289, 87975
        .AddLineSegment -786289, 82989
        .AddLineSegment -783250, 82989
        .AddLineSegment -783250, 81964
        .AddLineSegment -787517, 81964
        .CloseSubPath

        .StartSubPath
        .AddLineSegment -778501, 84199
        .AddLineSegment -777316, 83835
        .AddBezierSegment -777316, 83835, -777498, 83168, -777802, 82672, -778225, 82348
        .AddBezierSegment -778225, 82348, -778648, 82025, -779184, 81862, -779835, 81862
        .AddBezierSegment -779835, 81862, -780641, 81862, -781302, 82138, -781822, 82688
        .AddBezierSegment -781822, 82688, -782340, 83238, -782599, 83990, -782599, 84946
        .AddBezierSegment -782599, 84946, -782599, 85955, -782338, 86739, -781817, 87298
        .AddBezierSegment -781817, 87298, -781297, 87857, -780611, 88136, -779761, 88136
        .AddBezierSegment -779761, 88136, -779019, 88136, -778417, 87916, -777954, 87477
        .AddBezierSegment -777954, 87477, -777677, 87218, -777471, 86845, -777333, 86358
        .AddLineSegment -778544, 86070
        .AddBezierSegment -778544, 86070, -778616, 86384, -778765, 86632, -778992, 86814
        .AddBezierSegment -778992, 86814, -779221, 86996, -779498, 87087, -779822, 87087
        .AddBezierSegment -779822, 87087, -780272, 87087, -780638, 86926, -780917, 86602
        .AddBezierSegment -780917, 86602, -781198, 86281, -781338, 85759, -781338, 85036
        .AddBezierSegment -781338, 85036, -781338, 84271, -781199, 83725, -780924, 83401
        .AddBezierSegment -780924, 83401, -780648, 83075, -780289, 82912, -779848, 82912
        .AddBezierSegment -779848, 82912, -779522, 82912, -779241, 83015, -779007, 83223
        .AddBezierSegment -779007, 83223, -778772, 83430, -778603, 83755, -778501, 84199
        .CloseSubPath

        Set .brush = New waBrush
        .brush.SetColor RGB(43, 42, 41), 100
    End With
    MyCurves.Add path

    Set path = New waPath
    With path
        .StartSubPath
        .AddLineSegment -1670499, 119627
        .AddLineSegment -1732361, 115076
        .AddBezierSegment -1732361, 115076, -1762635, 89710, -1757038, -9203, -1754782, -50659
        .AddLineSegment -1670499, -57346
        .AddLineSegment -1670499, 119627
        .CloseSubPath

        Set .brush = New waBrush
        .brush.SetLinearGradient -1714000, 59000, -1714000, 236000, RGB(228, 239, 240), RGB(177, 221, 227), WrapModeTileFlipXY
    End With
    MyCurves.Add path

    Set path = New waPath
    With path
        .StartSubPath
        .AddLineSegment 70499, 119627
        .AddLineSegment 132361, 115076
        .AddBezierSegment 132361, 115076, 162635, 89710, 157038, -9203, 154782, -50659
        .AddLineSegment 70499, -57346
        .AddLineSegment 70499, 119627
        .CloseSubPath

        Set .brush = New waBrush
        .brush.SetLinearGradient 114000, 59000, 114000, 236000, RGB(228, 239, 240), RGB(177, 221, 227), WrapModeTileFlipXY
    End With
    MyCurves.Add path

    Set path = New waPath
    With path
        .StartSubPath
        .AddLineSegment -1617082, 104791
        .AddLineSegment -1612534, 104966
        .AddBezierSegment -1612534, 104966, -1612260, 104374, -1611712, 103942, -1610879, 103663
        .AddBezierSegment -1610879, 103663, -1610046, 103385, -1608923, 103247, -1607509, 103247
        .AddBezierSegment -1607509, 103247, -1606015, 103247, -1604886, 103370, -1604129, 103617
        .AddBezierSegment -1604129, 103617, -1603371, 103864, -1602989, 104153, -1602989, 104483
        .AddBezierSegment -1602989, 104483, -1602989, 104697, -1603150, 104877, -1603467, 105027
        .AddBezierSegment -1603467, 105027, -1603784, 105176, -1604343, 105305, -1605139, 105414
        .AddBezierSegment -1605139, 105414, -1605682, 105490, -1606918, 105619, -1608853, 105808
        .AddBezierSegment -1608853, 105808, -1611341, 106050, -1613088, 106347, -1614093, 106701
        .AddBezierSegment -1614093, 106701, -1615506, 107197, -1616211, 107801, -1616211, 108514
        .AddBezierSegment -1616211, 108514, -1616211, 108973, -1615878, 109403, -1615216, 109803
        .AddBezierSegment -1615216, 109803, -1614550, 110203, -1613593, 110506, -1612341, 110715
        .AddBezierSegment -1612341, 110715, -1611089, 110924, -1609584, 111029, -1607810, 111029
        .AddBezierSegment -1607810, 111029, -1604924, 111029, -1602747, 110780, -1601291, 110280
        .AddBezierSegment -1601291, 110280, -1599829, 109780, -1599066, 109115, -1598990, 108280
        .AddLineSegment -1603666, 108217
        .AddBezierSegment -1603666, 108217, -1603865, 108682, -1604290, 109014, -1604940, 109217
        .AddBezierSegment -1604940, 109217, -1605590, 109420, -1606569, 109523, -1607869, 109523
        .AddBezierSegment -1607869, 109523, -1609213, 109523, -1610266, 109414, -1611024, 109196
        .AddBezierSegment -1611024, 109196, -1611513, 109056, -1611761, 108870, -1611761, 108636
        .AddBezierSegment -1611761, 108636, -1611761, 108422, -1611529, 108240, -1611067, 108090
        .AddBezierSegment -1611067, 108090, -1610481, 107895, -1609057, 107694, -1606794, 107485
        .AddBezierSegment -1606794, 107485, -1604532, 107278, -1602860, 107061, -1601780, 106839
        .AddBezierSegment -1601780, 106839, -1600694, 106615, -1599845, 106312, -1599232, 105925
        .AddBezierSegment -1599232, 105925, -1598619, 105538, -1598313, 105061, -1598313, 104494
        .AddBezierSegment -1598313, 104494, -1598313, 103979, -1598679, 103496, -1599415, 103046
        .AddBezierSegment -1599415, 103046, -1600151, 102596, -1601188, 102262, -1602532, 102044
        .AddBezierSegment -1602532, 102044, -1603876, 101824, -1605553, 101716, -1607558, 101716
        .AddBezierSegment -1607558, 101716, -1610476, 101716, -1612717, 101977, -1614281, 102502
        .AddBezierSegment -1614281, 102502, -1615845, 103025, -1616775, 103789, -1617082, 104791
        .CloseSubPath

        .StartSubPath
        .AddLineSegment -1589155, 101879
        .AddLineSegment -1589155, 109347
        .AddLineSegment -1595991, 109347
        .AddLineSegment -1595991, 110866
        .AddLineSegment -1577674, 110866
        .AddLineSegment -1577674, 109347
        .AddLineSegment -1584479, 109347
        .AddLineSegment -1584479, 101879
        .AddLineSegment -1589155, 101879
        .CloseSubPath

        .StartSubPath
        .AddLineSegment -1556240, 101879
        .AddLineSegment -1561276, 101879
        .AddLineSegment -1563270, 103925
        .AddLineSegment -1572531, 103925
        .AddLineSegment -1574449, 101879
        .AddLineSegment -1579394, 101879
        .AddLineSegment -1570429, 110866
        .AddLineSegment -1565490, 110866
        .AddLineSegment -1556240, 101879
        .CloseSubPath

        .StartSubPath
        .AddLineSegment -1564786, 105444
        .AddLineSegment -1567995, 108770
        .AddLineSegment -1571112, 105444
        .AddLineSegment -1564786, 105444
        .CloseSubPath

        .StartSubPath
        .AddLineSegment -1553752, 101879
        .AddLineSegment -1553752, 110866
        .AddLineSegment -1549226, 110866
        .AddLineSegment -1539724, 104843
        .AddLineSegment -1539724, 110866
        .AddLineSegment -1535402, 110866
        .AddLineSegment -1535402, 101879
        .AddLineSegment -1540073, 101879
        .AddLineSegment -1549430, 107787
        .AddLineSegment -1549430, 101879
        .AddLineSegment -1553752, 101879
        .CloseSubPath

        .StartSubPath
        .AddLineSegment -1530463, 110866
        .AddLineSegment -1521961, 110866
        .AddBezierSegment -1521961, 110866, -1520042, 110866, -1518580, 110809, -1517575, 110694
        .AddBezierSegment -1517575, 110694, -1516226, 110540, -1515065, 110263, -1514103, 109866
        .AddBezierSegment -1514103, 109866, -1513135, 109471, -1512404, 108985, -1511899, 108410
        .AddBezierSegment -1511899, 108410, -1511399, 107837, -1511147, 107127, -1511147, 106287
        .AddBezierSegment -1511147, 106287, -1511147, 105546, -1511383, 104910, -1511856, 104374
        .AddBezierSegment -1511856, 104374, -1512431, 103720, -1513254, 103191, -1514323, 102787
        .AddBezierSegment -1514323, 102787, -1515129, 102479, -1516220, 102241, -1517591, 102069
        .AddBezierSegment -1517591, 102069, -1518618, 101942, -1519988, 101879, -1521708, 101879
        .AddLineSegment -1530463, 101879
        .AddLineSegment -1530463, 110866
        .CloseSubPath

        .StartSubPath
        .AddLineSegment -1525787, 109347
        .AddLineSegment -1525787, 103398
        .AddLineSegment -1522321, 103398
        .AddBezierSegment -1522321, 103398, -1521025, 103398, -1520090, 103427, -1519515, 103483
        .AddBezierSegment -1519515, 103483, -1518757, 103556, -1518134, 103682, -1517639, 103858
        .AddBezierSegment -1517639, 103858, -1517139, 104034, -1516736, 104322, -1516425, 104724
        .AddBezierSegment -1516425, 104724, -1516107, 105126, -1515952, 105676, -1515952, 106370
        .AddBezierSegment -1515952, 106370, -1515952, 107065, -1516107, 107596, -1516425, 107969
        .AddBezierSegment -1516425, 107969, -1516736, 108341, -1517177, 108629, -1517741, 108839
        .AddBezierSegment -1517741, 108839, -1518306, 109046, -1519021, 109188, -1519891, 109261
        .AddBezierSegment -1519891, 109261, -1520536, 109318, -1521810, 109347, -1523702, 109347
        .AddLineSegment -1525787, 109347
        .CloseSubPath

        .StartSubPath
        .AddLineSegment -1486375, 101879
        .AddLineSegment -1491411, 101879
        .AddLineSegment -1493405, 103925
        .AddLineSegment -1502666, 103925
        .AddLineSegment -1504584, 101879
        .AddLineSegment -1509529, 101879
        .AddLineSegment -1500564, 110866
        .AddLineSegment -1495625, 110866
        .AddLineSegment -1486375, 101879
        .CloseSubPath

        .StartSubPath
        .AddLineSegment -1494921, 105444
        .AddLineSegment -1498129, 108770
        .AddLineSegment -1501246, 105444
        .AddLineSegment -1494921, 105444
        .CloseSubPath

        .StartSubPath
        .AddLineSegment -1483886, 101879
        .AddLineSegment -1483886, 110866
        .AddLineSegment -1474067, 110866
        .AddBezierSegment -1474067, 110866, -1471600, 110866, -1469805, 110784, -1468687, 110623
        .AddBezierSegment -1468687, 110623, -1467564, 110462, -1466671, 110175, -1465999, 109763
        .AddBezierSegment -1465999, 109763, -1465328, 109349, -1464989, 108879, -1464989, 108347
        .AddBezierSegment -1464989, 108347, -1464989, 107671, -1465494, 107115, -1466510, 106676
        .AddBezierSegment -1466510, 106676, -1467526, 106236, -1469042, 105960, -1471062, 105845
        .AddBezierSegment -1471062, 105845, -1470057, 105617, -1469230, 105366, -1468580, 105092
        .AddBezierSegment -1468580, 105092, -1467924, 104818, -1467042, 104331, -1465935, 103632
        .AddLineSegment -1463119, 101879
        .AddLineSegment -1468692, 101879
        .AddLineSegment -1472035, 103841
        .AddBezierSegment -1472035, 103841, -1473234, 104538, -1474056, 104977, -1474502, 105159
        .AddBezierSegment -1474502, 105159, -1474943, 105339, -1475410, 105464, -1475905, 105531
        .AddBezierSegment -1475905, 105531, -1476405, 105598, -1477190, 105632, -1478264, 105632
        .AddLineSegment -1479210, 105632
        .AddLineSegment -1479210, 101879
        .AddLineSegment -1483886, 101879
        .CloseSubPath

        .StartSubPath
        .AddLineSegment -1479210, 107063
        .AddLineSegment -1475760, 107063
        .AddBezierSegment -1475760, 107063, -1473524, 107063, -1472132, 107100, -1471573, 107174
        .AddBezierSegment -1471573, 107174, -1471020, 107247, -1470579, 107374, -1470267, 107554
        .AddBezierSegment -1470267, 107554, -1469950, 107734, -1469794, 107960, -1469794, 108230
        .AddBezierSegment -1469794, 108230, -1469794, 108533, -1470004, 108776, -1470418, 108964
        .AddBezierSegment -1470418, 108964, -1470831, 109150, -1471417, 109268, -1472170, 109316
        .AddBezierSegment -1472170, 109316, -1472551, 109337, -1473685, 109347, -1475572, 109347
        .AddLineSegment -1479210, 109347
        .AddLineSegment -1479210, 107063
        .CloseSubPath

        .StartSubPath
        .AddLineSegment -1460598, 110866
        .AddLineSegment -1452095, 110866
        .AddBezierSegment -1452095, 110866, -1450176, 110866, -1448715, 110809, -1447709, 110694
        .AddBezierSegment -1447709, 110694, -1446361, 110540, -1445200, 110263, -1444237, 109866
        .AddBezierSegment -1444237, 109866, -1443270, 109471, -1442539, 108985, -1442034, 108410
        .AddBezierSegment -1442034, 108410, -1441534, 107837, -1441282, 107127, -1441282, 106287
        .AddBezierSegment -1441282, 106287, -1441282, 105546, -1441518, 104910, -1441991, 104374
        .AddBezierSegment -1441991, 104374, -1442566, 103720, -1443388, 103191, -1444458, 102787
        .AddBezierSegment -1444458, 102787, -1445264, 102479, -1446355, 102241, -1447726, 102069
        .AddBezierSegment -1447726, 102069, -1448752, 101942, -1450123, 101879, -1451843, 101879
        .AddLineSegment -1460598, 101879
        .AddLineSegment -1460598, 110866
        .CloseSubPath

        .StartSubPath
        .AddLineSegment -1455922, 109347
        .AddLineSegment -1455922, 103398
        .AddLineSegment -1452455, 103398
        .AddBezierSegment -1452455, 103398, -1451160, 103398, -1450225, 103427, -1449650, 103483
        .AddBezierSegment -1449650, 103483, -1448892, 103556, -1448268, 103682, -1447774, 103858
        .AddBezierSegment -1447774, 103858, -1447274, 104034, -1446871, 104322, -1446560, 104724
        .AddBezierSegment -1446560, 104724, -1446242, 105126, -1446086, 105676, -1446086, 106370
        .AddBezierSegment -1446086, 106370, -1446086, 107065, -1446242, 107596, -1446560, 107969
        .AddBezierSegment -1446560, 107969, -1446871, 108341, -1447312, 108629, -1447876, 108839
        .AddBezierSegment -1447876, 108839, -1448440, 109046, -1449155, 109188, -1450026, 109261
        .AddBezierSegment -1450026, 109261, -1450671, 109318, -1451945, 109347, -1453837, 109347
        .AddLineSegment -1455922, 109347
        .CloseSubPath

        .StartSubPath
        .AddLineSegment -1428350, 101879
        .AddLineSegment -1428350, 110866
        .AddLineSegment -1420874, 110866
        .AddBezierSegment -1420874, 110866, -1418041, 110866, -1416198, 110822, -1415338, 110732
        .AddBezierSegment -1415338, 110732, -1414016, 110596, -1412909, 110303, -1412016, 109851
        .AddBezierSegment -1412016, 109851, -1411124, 109399, -1410678, 108816, -1410678, 108100
        .AddBezierSegment -1410678, 108100, -1410678, 107548, -1410936, 107086, -1411447, 106709
        .AddBezierSegment -1411447, 106709, -1411963, 106333, -1412613, 106038, -1413403, 105822
        .AddBezierSegment -1413403, 105822, -1414193, 105609, -1414999, 105466, -1415816, 105397
        .AddBezierSegment -1415816, 105397, -1416929, 105312, -1418536, 105268, -1420643, 105268
        .AddLineSegment -1423674, 105268
        .AddLineSegment -1423674, 101879
        .AddLineSegment -1428350, 101879
        .CloseSubPath

        .StartSubPath
        .AddLineSegment -1423674, 109347
        .AddLineSegment -1423674, 106787
        .AddLineSegment -1421132, 106787
        .AddBezierSegment -1421132, 106787, -1419299, 106787, -1418079, 106835, -1417461, 106929
        .AddBezierSegment -1417461, 106929, -1416843, 107023, -1416359, 107169, -1416010, 107372
        .AddBezierSegment -1416010, 107372, -1415660, 107573, -1415483, 107805, -1415483, 108073
        .AddBezierSegment -1415483, 108073, -1415483, 108401, -1415730, 108671, -1416219, 108885
        .AddBezierSegment -1416219, 108885, -1416714, 109098, -1417332, 109232, -1418090, 109286
        .AddBezierSegment -1418090, 109286, -1418643, 109326, -1419756, 109347, -1421433, 109347
        .AddLineSegment -1423674, 109347
        .CloseSubPath

        .StartSubPath
        .AddLineSegment -1406711, 101879
        .AddLineSegment -1406711, 110790
        .AddLineSegment -1402035, 110790
        .AddLineSegment -1402035, 103398
        .AddLineSegment -1390459, 103398
        .AddLineSegment -1390459, 101879
        .AddLineSegment -1406711, 101879
        .CloseSubPath

        .StartSubPath
        .AddLineSegment -1388110, 106318
        .AddBezierSegment -1388110, 106318, -1388110, 107234, -1387755, 108004, -1387051, 108627
        .AddBezierSegment -1387051, 108627, -1386530, 109086, -1385809, 109498, -1384901, 109862
        .AddBezierSegment -1384901, 109862, -1383993, 110226, -1382998, 110496, -1381918, 110673
        .AddBezierSegment -1381918, 110673, -1380478, 110910, -1378817, 111029, -1376936, 111029
        .AddBezierSegment -1376936, 111029, -1373528, 111029, -1370809, 110617, -1368761, 109795
        .AddBezierSegment -1368761, 109795, -1366718, 108971, -1365697, 107826, -1365697, 106360
        .AddBezierSegment -1365697, 106360, -1365697, 104906, -1366713, 103768, -1368739, 102948
        .AddBezierSegment -1368739, 102948, -1370766, 102126, -1373480, 101716, -1376871, 101716
        .AddBezierSegment -1376871, 101716, -1380306, 101716, -1383042, 102124, -1385068, 102942
        .AddBezierSegment -1385068, 102942, -1387094, 103757, -1388110, 104883, -1388110, 106318
        .CloseSubPath

        .StartSubPath
        .AddLineSegment -1383305, 106379
        .AddBezierSegment -1383305, 106379, -1383305, 105360, -1382697, 104588, -1381488, 104061
        .AddBezierSegment -1381488, 104061, -1380279, 103536, -1378742, 103272, -1376882, 103272
        .AddBezierSegment -1376882, 103272, -1375017, 103272, -1373491, 103534, -1372292, 104054
        .AddBezierSegment -1372292, 104054, -1371099, 104577, -1370502, 105360, -1370502, 106404
        .AddBezierSegment -1370502, 106404, -1370502, 107435, -1371083, 108205, -1372249, 108711
        .AddBezierSegment -1372249, 108711, -1373410, 109219, -1374953, 109473, -1376882, 109473
        .AddBezierSegment -1376882, 109473, -1378806, 109473, -1380360, 109215, -1381536, 108703
        .AddBezierSegment -1381536, 108703, -1382714, 108188, -1383305, 107414, -1383305, 106379
        .CloseSubPath

        .StartSubPath
        .AddLineSegment -1356867, 101879
        .AddLineSegment -1356867, 109347
        .AddLineSegment -1363704, 109347
        .AddLineSegment -1363704, 110866
        .AddLineSegment -1345387, 110866
        .AddLineSegment -1345387, 109347
        .AddLineSegment -1352191, 109347
        .AddLineSegment -1352191, 101879
        .AddLineSegment -1356867, 101879
        .CloseSubPath

        .StartSubPath
        .AddLineSegment -1337169, 101879
        .AddLineSegment -1337169, 109347
        .AddLineSegment -1344005, 109347
        .AddLineSegment -1344005, 110866
        .AddLineSegment -1325688, 110866
        .AddLineSegment -1325688, 109347
        .AddLineSegment -1332493, 109347
        .AddLineSegment -1332493, 101879
        .AddLineSegment -1337169, 101879
        .CloseSubPath

        .StartSubPath
        .AddLineSegment -1322662, 101879
        .AddLineSegment -1322662, 110866
        .AddLineSegment -1305571, 110866
        .AddLineSegment -1305571, 109347
        .AddLineSegment -1317986, 109347
        .AddLineSegment -1317986, 107351
        .AddLineSegment -1306442, 107351
        .AddLineSegment -1306442, 105833
        .AddLineSegment -1317986, 105833
        .AddLineSegment -1317986, 103398
        .AddLineSegment -1305120, 103398
        .AddLineSegment -1305120, 101879
        .AddLineSegment -1322662, 101879
        .CloseSubPath

        .StartSubPath
        .AddLineSegment -1301153, 101879
        .AddLineSegment -1301153, 110866
        .AddLineSegment -1291333, 110866
        .AddBezierSegment -1291333, 110866, -1288866, 110866, -1287071, 110784, -1285953, 110623
        .AddBezierSegment -1285953, 110623, -1284830, 110462, -1283938, 110175, -1283266, 109763
        .AddBezierSegment -1283266, 109763, -1282594, 109349, -1282255, 108879, -1282255, 108347
        .AddBezierSegment -1282255, 108347, -1282255, 107671, -1282761, 107115, -1283777, 106676
        .AddBezierSegment -1283777, 106676, -1284792, 106236, -1286308, 105960, -1288329, 105845
        .AddBezierSegment -1288329, 105845, -1287324, 105617, -1286496, 105366, -1285846, 105092
        .AddBezierSegment -1285846, 105092, -1285190, 104818, -1284309, 104331, -1283201, 103632
        .AddLineSegment -1280385, 101879
        .AddLineSegment -1285959, 101879
        .AddLineSegment -1289302, 103841
        .AddBezierSegment -1289302, 103841, -1290500, 104538, -1291323, 104977, -1291769, 105159
        .AddBezierSegment -1291769, 105159, -1292209, 105339, -1292677, 105464, -1293172, 105531
        .AddBezierSegment -1293172, 105531, -1293671, 105598, -1294456, 105632, -1295531, 105632
        .AddLineSegment -1296477, 105632
        .AddLineSegment -1296477, 101879
        .AddLineSegment -1301153, 101879
        .CloseSubPath

        .StartSubPath
        .AddLineSegment -1296477, 107063
        .AddLineSegment -1293026, 107063
        .AddBezierSegment -1293026, 107063, -1290791, 107063, -1289399, 107100, -1288840, 107174
        .AddBezierSegment -1288840, 107174, -1288286, 107247, -1287845, 107374, -1287534, 107554
        .AddBezierSegment -1287534, 107554, -1287216, 107734, -1287061, 107960, -1287061, 108230
        .AddBezierSegment -1287061, 108230, -1287061, 108533, -1287270, 108776, -1287684, 108964
        .AddBezierSegment -1287684, 108964, -1288098, 109150, -1288684, 109268, -1289436, 109316
        .AddBezierSegment -1289436, 109316, -1289818, 109337, -1290952, 109347, -1292838, 109347
        .AddLineSegment -1296477, 109347
        .AddLineSegment -1296477, 107063
        .CloseSubPath

        .StartSubPath
        .AddLineSegment -1258586, 101879
        .AddLineSegment -1263004, 101879
        .AddLineSegment -1263004, 108385
        .AddBezierSegment -1263004, 108385, -1264616, 107795, -1266519, 107358, -1268711, 107075
        .AddLineSegment -1268711, 108644
        .AddBezierSegment -1268711, 108644, -1267556, 108791, -1266303, 109069, -1264954, 109481
        .AddBezierSegment -1264954, 109481, -1263600, 109891, -1262676, 110370, -1262170, 110916
        .AddLineSegment -1258586, 110916
        .AddLineSegment -1258586, 101879
        .CloseSubPath

        .StartSubPath
        .AddLineSegment -1236974, 108665
        .AddLineSegment -1241263, 108481
        .AddBezierSegment -1241263, 108481, -1241365, 108824, -1241639, 109077, -1242080, 109240
        .AddBezierSegment -1242080, 109240, -1242521, 109403, -1243090, 109485, -1243795, 109485
        .AddBezierSegment -1243795, 109485, -1244730, 109485, -1245520, 109322, -1246165, 108996
        .AddBezierSegment -1246165, 108996, -1246810, 108669, -1247213, 107989, -1247379, 106956
        .AddBezierSegment -1247379, 106956, -1246278, 107462, -1244912, 107715, -1243273, 107715
        .AddBezierSegment -1243273, 107715, -1241430, 107715, -1239844, 107441, -1238527, 106893
        .AddBezierSegment -1238527, 106893, -1237210, 106343, -1236555, 105636, -1236555, 104768
        .AddBezierSegment -1236555, 104768, -1236555, 103847, -1237248, 103107, -1238635, 102550
        .AddBezierSegment -1238635, 102550, -1240022, 101994, -1241801, 101716, -1243972, 101716
        .AddBezierSegment -1243972, 101716, -1246299, 101716, -1248218, 102069, -1249717, 102776
        .AddBezierSegment -1249717, 102776, -1251217, 103483, -1251969, 104640, -1251969, 106251
        .AddBezierSegment -1251969, 106251, -1251969, 107901, -1251185, 109092, -1249621, 109822
        .AddBezierSegment -1249621, 109822, -1248057, 110552, -1246025, 110916, -1243526, 110916
        .AddBezierSegment -1243526, 110916, -1241774, 110916, -1240323, 110724, -1239172, 110339
        .AddBezierSegment -1239172, 110339, -1238022, 109956, -1237291, 109397, -1236974, 108665
        .CloseSubPath

        .StartSubPath
        .AddLineSegment -1247003, 104904
        .AddBezierSegment -1247003, 104904, -1247003, 104343, -1246670, 103910, -1246009, 103605
        .AddBezierSegment -1246009, 103605, -1245348, 103299, -1244595, 103147, -1243741, 103147
        .AddBezierSegment -1243741, 103147, -1242924, 103147, -1242241, 103272, -1241693, 103521
        .AddBezierSegment -1241693, 103521, -1241150, 103772, -1240876, 104182, -1240876, 104751
        .AddBezierSegment -1240876, 104751, -1240876, 105337, -1241172, 105766, -1241757, 106040
        .AddBezierSegment -1241757, 106040, -1242349, 106312, -1243080, 106448, -1243961, 106448
        .AddBezierSegment -1243961, 106448, -1244816, 106448, -1245531, 106318, -1246122, 106056
        .AddBezierSegment -1246122, 106056, -1246708, 105797, -1247003, 105412, -1247003, 104904
        .CloseSubPath

        .StartSubPath
        .AddLineSegment -1226537, 110916
        .AddBezierSegment -1226537, 110916, -1224301, 110916, -1222554, 110604, -1221291, 109981
        .AddBezierSegment -1221291, 109981, -1219792, 109245, -1219039, 108021, -1219039, 106314
        .AddBezierSegment -1219039, 106314, -1219039, 104609, -1219797, 103383, -1221307, 102638
        .AddBezierSegment -1221307, 102638, -1222559, 102023, -1224301, 101716, -1226537, 101716
        .AddBezierSegment -1226537, 101716, -1228783, 101716, -1230594, 102052, -1231970, 102726
        .AddBezierSegment -1231970, 102726, -1233346, 103400, -1234034, 104603, -1234034, 106331
        .AddBezierSegment -1234034, 106331, -1234034, 108027, -1233276, 109249, -1231766, 109993
        .AddBezierSegment -1231766, 109993, -1230514, 110609, -1228772, 110916, -1226537, 110916
        .CloseSubPath

        .StartSubPath
        .AddLineSegment -1226537, 109485
        .AddBezierSegment -1226537, 109485, -1227063, 109485, -1227536, 109418, -1227945, 109284
        .AddBezierSegment -1227945, 109284, -1228359, 109153, -1228681, 108914, -1228907, 108569
        .AddBezierSegment -1228907, 108569, -1229208, 108123, -1229358, 107370, -1229358, 106314
        .AddBezierSegment -1229358, 106314, -1229358, 105255, -1229224, 104529, -1228955, 104134
        .AddBezierSegment -1228955, 104134, -1228686, 103739, -1228348, 103475, -1227939, 103343
        .AddBezierSegment -1227939, 103343, -1227531, 103211, -1227063, 103147, -1226537, 103147
        .AddBezierSegment -1226537, 103147, -1226010, 103147, -1225537, 103213, -1225123, 103347
        .AddBezierSegment -1225123, 103347, -1224715, 103479, -1224392, 103718, -1224166, 104063
        .AddBezierSegment -1224166, 104063, -1223866, 104504, -1223715, 105255, -1223715, 106314
        .AddBezierSegment -1223715, 106314, -1223715, 107370, -1223849, 108096, -1224118, 108492
        .AddBezierSegment -1224118, 108492, -1224387, 108887, -1224725, 109153, -1225134, 109284
        .AddBezierSegment -1225134, 109284, -1225542, 109418, -1226010, 109485, -1226537, 109485
        .CloseSubPath

        .StartSubPath
        .AddLineSegment -1208601, 110916
        .AddBezierSegment -1208601, 110916, -1206366, 110916, -1204619, 110604, -1203356, 109981
        .AddBezierSegment -1203356, 109981, -1201856, 109245, -1201104, 108021, -1201104, 106314
        .AddBezierSegment -1201104, 106314, -1201104, 104609, -1201862, 103383, -1203372, 102638
        .AddBezierSegment -1203372, 102638, -1204624, 102023, -1206366, 101716, -1208601, 101716
        .AddBezierSegment -1208601, 101716, -1210848, 101716, -1212659, 102052, -1214035, 102726
        .AddBezierSegment -1214035, 102726, -1215411, 103400, -1216099, 104603, -1216099, 106331
        .AddBezierSegment -1216099, 106331, -1216099, 108027, -1215341, 109249, -1213831, 109993
        .AddBezierSegment -1213831, 109993, -1212579, 110609, -1210837, 110916, -1208601, 110916
        .CloseSubPath

        .StartSubPath
        .AddLineSegment -1208601, 109485
        .AddBezierSegment -1208601, 109485, -1209128, 109485, -1209601, 109418, -1210009, 109284
        .AddBezierSegment -1210009, 109284, -1210423, 109153, -1210746, 108914, -1210972, 108569
        .AddBezierSegment -1210972, 108569, -1211273, 108123, -1211423, 107370, -1211423, 106314
        .AddBezierSegment -1211423, 106314, -1211423, 105255, -1211289, 104529, -1211020, 104134
        .AddBezierSegment -1211020, 104134, -1210751, 103739, -1210412, 103475, -1210004, 103343
        .AddBezierSegment -1210004, 103343, -1209596, 103211, -1209128, 103147, -1208601, 103147
        .AddBezierSegment -1208601, 103147, -1208074, 103147, -1207602, 103213, -1207188, 103347
        .AddBezierSegment -1207188, 103347, -1206779, 103479, -1206457, 103718, -1206231, 104063
        .AddBezierSegment -1206231, 104063, -1205930, 104504, -1205780, 105255, -1205780, 106314
        .AddBezierSegment -1205780, 106314, -1205780, 107370, -1205914, 108096, -1206183, 108492
        .AddBezierSegment -1206183, 108492, -1206452, 108887, -1206790, 109153, -1207199, 109284
        .AddBezierSegment -1207199, 109284, -1207607, 109418, -1208074, 109485, -1208601, 109485
        .CloseSubPath

        Set .brush = New waBrush
        .brush.SetLinearGradient -1409000, 119000, -1409000, 104000, RGB(43, 42, 41), RGB(89, 89, 89), WrapModeTileFlipXY
    End With
    MyCurves.Add path
    InitStandardPlotter1600_2Dmodel_3 MyCurves
End Sub

' -----------------------------------------------------------------------------
' InitStandardPlotter1600_2Dmodel_3 (Sub), source lines 8382-9677
' -----------------------------------------------------------------------------
Private Sub InitStandardPlotter1600_2Dmodel_3(MyCurves As Collection)
    Dim path As waPath
    Set path = New waPath
    With path
        .StartSubPath
        .AddLineSegment -1243886, 0
        .AddLineSegment -1215734, 0
        .AddLineSegment -1215734, 6430
        .AddLineSegment -1219438, 6430
        .AddLineSegment -1219438, 22145
        .AddLineSegment -1240181, 22145
        .AddLineSegment -1240181, 6430
        .AddLineSegment -1243886, 6430
        .AddLineSegment -1243886, 0
        .CloseSubPath

        Set .brush = New waBrush
        .brush.SetColor RGB(43, 42, 41), 100
    End With
    MyCurves.Add path

    Set path = New waPath
    With path
        .StartSubPath
        .AddLineSegment -370125, 0
        .AddLineSegment -341974, 0
        .AddLineSegment -341974, 6430
        .AddLineSegment -345678, 6430
        .AddLineSegment -345678, 22145
        .AddLineSegment -366421, 22145
        .AddLineSegment -366421, 6430
        .AddLineSegment -370125, 6430
        .AddLineSegment -370125, 0
        .CloseSubPath

        Set .brush = New waBrush
        .brush.SetColor RGB(43, 42, 41), 100
    End With
    MyCurves.Add path

    Set path = New waPath
    With path
        .StartSubPath
        .AddLineSegment -1625000, 81964
        .AddLineSegment -1625000, 88026
        .AddLineSegment -1623037, 88026
        .AddBezierSegment -1623037, 88026, -1622293, 88026, -1621809, 87997, -1621584, 87936
        .AddBezierSegment -1621584, 87936, -1621236, 87844, -1620946, 87647, -1620711, 87342
        .AddBezierSegment -1620711, 87342, -1620477, 87037, -1620360, 86643, -1620360, 86161
        .AddBezierSegment -1620360, 86161, -1620360, 85788, -1620428, 85476, -1620562, 85222
        .AddBezierSegment -1620562, 85222, -1620697, 84968, -1620868, 84769, -1621076, 84624
        .AddBezierSegment -1621076, 84624, -1621283, 84480, -1621495, 84384, -1621709, 84338
        .AddBezierSegment -1621709, 84338, -1622001, 84280, -1622423, 84250, -1622976, 84250
        .AddLineSegment -1623772, 84250
        .AddLineSegment -1623772, 81964
        .AddLineSegment -1625000, 81964
        .CloseSubPath

        .StartSubPath
        .AddLineSegment -1623772, 87002
        .AddLineSegment -1623772, 85275
        .AddLineSegment -1623105, 85275
        .AddBezierSegment -1623105, 85275, -1622624, 85275, -1622303, 85307, -1622141, 85371
        .AddBezierSegment -1622141, 85371, -1621979, 85434, -1621852, 85533, -1621760, 85670
        .AddBezierSegment -1621760, 85670, -1621668, 85805, -1621622, 85962, -1621622, 86142
        .AddBezierSegment -1621622, 86142, -1621622, 86364, -1621687, 86546, -1621815, 86690
        .AddBezierSegment -1621815, 86690, -1621945, 86834, -1622107, 86924, -1622306, 86961
        .AddBezierSegment -1622306, 86961, -1622451, 86988, -1622744, 87002, -1623184, 87002
        .AddLineSegment -1623772, 87002
        .CloseSubPath

        .StartSubPath
        .AddLineSegment -1619319, 81964
        .AddLineSegment -1619319, 87975
        .AddLineSegment -1618091, 87975
        .AddLineSegment -1618091, 82989
        .AddLineSegment -1615052, 82989
        .AddLineSegment -1615052, 81964
        .AddLineSegment -1619319, 81964
        .CloseSubPath

        .StartSubPath
        .AddLineSegment -1614435, 84959
        .AddBezierSegment -1614435, 84959, -1614435, 85577, -1614342, 86096, -1614157, 86516
        .AddBezierSegment -1614157, 86516, -1614020, 86825, -1613831, 87103, -1613592, 87349
        .AddBezierSegment -1613592, 87349, -1613354, 87594, -1613093, 87776, -1612809, 87896
        .AddBezierSegment -1612809, 87896, -1612431, 88056, -1611995, 88136, -1611501, 88136
        .AddBezierSegment -1611501, 88136, -1610607, 88136, -1609892, 87858, -1609355, 87304
        .AddBezierSegment -1609355, 87304, -1608819, 86748, -1608551, 85976, -1608551, 84987
        .AddBezierSegment -1608551, 84987, -1608551, 84006, -1608817, 83238, -1609349, 82685
        .AddBezierSegment -1609349, 82685, -1609881, 82131, -1610594, 81854, -1611484, 81854
        .AddBezierSegment -1611484, 81854, -1612386, 81854, -1613104, 82129, -1613636, 82681
        .AddBezierSegment -1613636, 82681, -1614168, 83231, -1614435, 83990, -1614435, 84959
        .CloseSubPath

        .StartSubPath
        .AddLineSegment -1613173, 84999
        .AddBezierSegment -1613173, 84999, -1613173, 84312, -1613014, 83791, -1612696, 83436
        .AddBezierSegment -1612696, 83436, -1612379, 83082, -1611975, 82904, -1611487, 82904
        .AddBezierSegment -1611487, 82904, -1610997, 82904, -1610597, 83080, -1610282, 83432
        .AddBezierSegment -1610282, 83432, -1609969, 83784, -1609812, 84312, -1609812, 85016
        .AddBezierSegment -1609812, 85016, -1609812, 85712, -1609965, 86231, -1610271, 86573
        .AddBezierSegment -1610271, 86573, -1610575, 86916, -1610980, 87087, -1611487, 87087
        .AddBezierSegment -1611487, 87087, -1611992, 87087, -1612400, 86913, -1612709, 86567
        .AddBezierSegment -1612709, 86567, -1613018, 86220, -1613173, 85698, -1613173, 84999
        .CloseSubPath

        .StartSubPath
        .AddLineSegment -1606232, 81964
        .AddLineSegment -1606232, 87002
        .AddLineSegment -1608027, 87002
        .AddLineSegment -1608027, 88026
        .AddLineSegment -1603218, 88026
        .AddLineSegment -1603218, 87002
        .AddLineSegment -1605004, 87002
        .AddLineSegment -1605004, 81964
        .AddLineSegment -1606232, 81964
        .CloseSubPath

        .StartSubPath
        .AddLineSegment -1598546, 84199
        .AddLineSegment -1597360, 83835
        .AddBezierSegment -1597360, 83835, -1597543, 83168, -1597846, 82672, -1598269, 82348
        .AddBezierSegment -1598269, 82348, -1598693, 82025, -1599229, 81862, -1599879, 81862
        .AddBezierSegment -1599879, 81862, -1600685, 81862, -1601347, 82138, -1601866, 82688
        .AddBezierSegment -1601866, 82688, -1602384, 83238, -1602644, 83990, -1602644, 84946
        .AddBezierSegment -1602644, 84946, -1602644, 85955, -1602383, 86739, -1601862, 87298
        .AddBezierSegment -1601862, 87298, -1601341, 87857, -1600655, 88136, -1599806, 88136
        .AddBezierSegment -1599806, 88136, -1599064, 88136, -1598461, 87916, -1597998, 87477
        .AddBezierSegment -1597998, 87477, -1597722, 87218, -1597516, 86845, -1597377, 86358
        .AddLineSegment -1598588, 86070
        .AddBezierSegment -1598588, 86070, -1598660, 86384, -1598810, 86632, -1599037, 86814
        .AddBezierSegment -1599037, 86814, -1599265, 86996, -1599542, 87087, -1599867, 87087
        .AddBezierSegment -1599867, 87087, -1600317, 87087, -1600682, 86926, -1600962, 86602
        .AddBezierSegment -1600962, 86602, -1601242, 86281, -1601382, 85759, -1601382, 85036
        .AddBezierSegment -1601382, 85036, -1601382, 84271, -1601244, 83725, -1600969, 83401
        .AddBezierSegment -1600969, 83401, -1600692, 83075, -1600334, 82912, -1599892, 82912
        .AddBezierSegment -1599892, 82912, -1599566, 82912, -1599285, 83015, -1599051, 83223
        .AddBezierSegment -1599051, 83223, -1598817, 83430, -1598647, 83755, -1598546, 84199
        .CloseSubPath

        .StartSubPath
        .AddLineSegment -1590848, 81964
        .AddLineSegment -1592170, 81964
        .AddLineSegment -1592694, 83344
        .AddLineSegment -1595125, 83344
        .AddLineSegment -1595629, 81964
        .AddLineSegment -1596927, 81964
        .AddLineSegment -1594573, 88026
        .AddLineSegment -1593277, 88026
        .AddLineSegment -1590848, 81964
        .CloseSubPath

        .StartSubPath
        .AddLineSegment -1593092, 84369
        .AddLineSegment -1593934, 86612
        .AddLineSegment -1594753, 84369
        .AddLineSegment -1593092, 84369
        .CloseSubPath

        .StartSubPath
        .AddLineSegment -1590161, 81964
        .AddLineSegment -1590161, 87975
        .AddLineSegment -1588933, 87975
        .AddLineSegment -1588933, 82989
        .AddLineSegment -1585894, 82989
        .AddLineSegment -1585894, 81964
        .AddLineSegment -1590161, 81964
        .CloseSubPath

        .StartSubPath
        .AddLineSegment -1581145, 84199
        .AddLineSegment -1579960, 83835
        .AddBezierSegment -1579960, 83835, -1580142, 83168, -1580445, 82672, -1580869, 82348
        .AddBezierSegment -1580869, 82348, -1581292, 82025, -1581828, 81862, -1582479, 81862
        .AddBezierSegment -1582479, 81862, -1583284, 81862, -1583946, 82138, -1584466, 82688
        .AddBezierSegment -1584466, 82688, -1584984, 83238, -1585243, 83990, -1585243, 84946
        .AddBezierSegment -1585243, 84946, -1585243, 85955, -1584982, 86739, -1584461, 87298
        .AddBezierSegment -1584461, 87298, -1583941, 87857, -1583255, 88136, -1582405, 88136
        .AddBezierSegment -1582405, 88136, -1581663, 88136, -1581061, 87916, -1580598, 87477
        .AddBezierSegment -1580598, 87477, -1580321, 87218, -1580115, 86845, -1579977, 86358
        .AddLineSegment -1581188, 86070
        .AddBezierSegment -1581188, 86070, -1581260, 86384, -1581409, 86632, -1581636, 86814
        .AddBezierSegment -1581636, 86814, -1581865, 86996, -1582141, 87087, -1582466, 87087
        .AddBezierSegment -1582466, 87087, -1582916, 87087, -1583282, 86926, -1583561, 86602
        .AddBezierSegment -1583561, 86602, -1583842, 86281, -1583982, 85759, -1583982, 85036
        .AddBezierSegment -1583982, 85036, -1583982, 84271, -1583843, 83725, -1583568, 83401
        .AddBezierSegment -1583568, 83401, -1583292, 83075, -1582933, 82912, -1582492, 82912
        .AddBezierSegment -1582492, 82912, -1582166, 82912, -1581885, 83015, -1581650, 83223
        .AddBezierSegment -1581650, 83223, -1581416, 83430, -1581247, 83755, -1581145, 84199
        .CloseSubPath

        Set .brush = New waBrush
        .brush.SetColor RGB(43, 42, 41), 100
    End With
    MyCurves.Add path

    Set path = New waPath
    With path
        .StartSubPath
        .AddLineSegment -20040, 81964
        .AddLineSegment -20040, 88026
        .AddLineSegment -18077, 88026
        .AddBezierSegment -18077, 88026, -17334, 88026, -16849, 87997, -16624, 87936
        .AddBezierSegment -16624, 87936, -16277, 87844, -15986, 87647, -15752, 87342
        .AddBezierSegment -15752, 87342, -15517, 87037, -15400, 86643, -15400, 86161
        .AddBezierSegment -15400, 86161, -15400, 85788, -15468, 85476, -15602, 85222
        .AddBezierSegment -15602, 85222, -15738, 84968, -15908, 84769, -16116, 84624
        .AddBezierSegment -16116, 84624, -16323, 84480, -16535, 84384, -16749, 84338
        .AddBezierSegment -16749, 84338, -17041, 84280, -17463, 84250, -18017, 84250
        .AddLineSegment -18812, 84250
        .AddLineSegment -18812, 81964
        .AddLineSegment -20040, 81964
        .CloseSubPath

        .StartSubPath
        .AddLineSegment -18812, 87002
        .AddLineSegment -18812, 85275
        .AddLineSegment -18145, 85275
        .AddBezierSegment -18145, 85275, -17664, 85275, -17343, 85307, -17181, 85371
        .AddBezierSegment -17181, 85371, -17019, 85434, -16892, 85533, -16800, 85670
        .AddBezierSegment -16800, 85670, -16708, 85805, -16662, 85962, -16662, 86142
        .AddBezierSegment -16662, 86142, -16662, 86364, -16727, 86546, -16855, 86690
        .AddBezierSegment -16855, 86690, -16985, 86834, -17147, 86924, -17346, 86961
        .AddBezierSegment -17346, 86961, -17492, 86988, -17784, 87002, -18224, 87002
        .AddLineSegment -18812, 87002
        .CloseSubPath

        .StartSubPath
        .AddLineSegment -14359, 81964
        .AddLineSegment -14359, 87975
        .AddLineSegment -13131, 87975
        .AddLineSegment -13131, 82989
        .AddLineSegment -10092, 82989
        .AddLineSegment -10092, 81964
        .AddLineSegment -14359, 81964
        .CloseSubPath

        .StartSubPath
        .AddLineSegment -9475, 84959
        .AddBezierSegment -9475, 84959, -9475, 85577, -9382, 86096, -9197, 86516
        .AddBezierSegment -9197, 86516, -9060, 86825, -8871, 87103, -8633, 87349
        .AddBezierSegment -8633, 87349, -8394, 87594, -8133, 87776, -7849, 87896
        .AddBezierSegment -7849, 87896, -7471, 88056, -7035, 88136, -6541, 88136
        .AddBezierSegment -6541, 88136, -5647, 88136, -4933, 87858, -4395, 87304
        .AddBezierSegment -4395, 87304, -3859, 86748, -3591, 85976, -3591, 84987
        .AddBezierSegment -3591, 84987, -3591, 84006, -3857, 83238, -4389, 82685
        .AddBezierSegment -4389, 82685, -4921, 82131, -5634, 81854, -6524, 81854
        .AddBezierSegment -6524, 81854, -7426, 81854, -8144, 82129, -8676, 82681
        .AddBezierSegment -8676, 82681, -9208, 83231, -9475, 83990, -9475, 84959
        .CloseSubPath

        .StartSubPath
        .AddLineSegment -8213, 84999
        .AddBezierSegment -8213, 84999, -8213, 84312, -8054, 83791, -7737, 83436
        .AddBezierSegment -7737, 83436, -7419, 83082, -7016, 82904, -6527, 82904
        .AddBezierSegment -6527, 82904, -6038, 82904, -5637, 83080, -5322, 83432
        .AddBezierSegment -5322, 83432, -5009, 83784, -4852, 84312, -4852, 85016
        .AddBezierSegment -4852, 85016, -4852, 85712, -5005, 86231, -5311, 86573
        .AddBezierSegment -5311, 86573, -5616, 86916, -6021, 87087, -6527, 87087
        .AddBezierSegment -6527, 87087, -7032, 87087, -7440, 86913, -7749, 86567
        .AddBezierSegment -7749, 86567, -8058, 86220, -8213, 85698, -8213, 84999
        .CloseSubPath

        .StartSubPath
        .AddLineSegment -1272, 81964
        .AddLineSegment -1272, 87002
        .AddLineSegment -3067, 87002
        .AddLineSegment -3067, 88026
        .AddLineSegment 1742, 88026
        .AddLineSegment 1742, 87002
        .AddLineSegment -44, 87002
        .AddLineSegment -44, 81964
        .AddLineSegment -1272, 81964
        .CloseSubPath

        .StartSubPath
        .AddLineSegment 6414, 84199
        .AddLineSegment 7599, 83835
        .AddBezierSegment 7599, 83835, 7417, 83168, 7114, 82672, 6691, 82348
        .AddBezierSegment 6691, 82348, 6267, 82025, 5731, 81862, 5081, 81862
        .AddBezierSegment 5081, 81862, 4275, 81862, 3613, 82138, 3094, 82688
        .AddBezierSegment 3094, 82688, 2576, 83238, 2316, 83990, 2316, 84946
        .AddBezierSegment 2316, 84946, 2316, 85955, 2577, 86739, 3098, 87298
        .AddBezierSegment 3098, 87298, 3619, 87857, 4304, 88136, 5154, 88136
        .AddBezierSegment 5154, 88136, 5896, 88136, 6499, 87916, 6962, 87477
        .AddBezierSegment 6962, 87477, 7238, 87218, 7444, 86845, 7582, 86358
        .AddLineSegment 6372, 86070
        .AddBezierSegment 6372, 86070, 6300, 86384, 6150, 86632, 5923, 86814
        .AddBezierSegment 5923, 86814, 5694, 86996, 5418, 87087, 5093, 87087
        .AddBezierSegment 5093, 87087, 4643, 87087, 4278, 86926, 3998, 86602
        .AddBezierSegment 3998, 86602, 3717, 86281, 3578, 85759, 3578, 85036
        .AddBezierSegment 3578, 85036, 3578, 84271, 3716, 83725, 3991, 83401
        .AddBezierSegment 3991, 83401, 4268, 83075, 4626, 82912, 5068, 82912
        .AddBezierSegment 5068, 82912, 5394, 82912, 5675, 83015, 5909, 83223
        .AddBezierSegment 5909, 83223, 6143, 83430, 6312, 83755, 6414, 84199
        .CloseSubPath

        .StartSubPath
        .AddLineSegment 14112, 81964
        .AddLineSegment 12790, 81964
        .AddLineSegment 12266, 83344
        .AddLineSegment 9835, 83344
        .AddLineSegment 9331, 81964
        .AddLineSegment 8033, 81964
        .AddLineSegment 10386, 88026
        .AddLineSegment 11683, 88026
        .AddLineSegment 14112, 81964
        .CloseSubPath

        .StartSubPath
        .AddLineSegment 11868, 84369
        .AddLineSegment 11026, 86612
        .AddLineSegment 10207, 84369
        .AddLineSegment 11868, 84369
        .CloseSubPath

        .StartSubPath
        .AddLineSegment 14799, 81964
        .AddLineSegment 14799, 87975
        .AddLineSegment 16027, 87975
        .AddLineSegment 16027, 82989
        .AddLineSegment 19066, 82989
        .AddLineSegment 19066, 81964
        .AddLineSegment 14799, 81964
        .CloseSubPath

        .StartSubPath
        .AddLineSegment 23815, 84199
        .AddLineSegment 25000, 83835
        .AddBezierSegment 25000, 83835, 24818, 83168, 24515, 82672, 24091, 82348
        .AddBezierSegment 24091, 82348, 23668, 82025, 23132, 81862, 22481, 81862
        .AddBezierSegment 22481, 81862, 21675, 81862, 21014, 82138, 20494, 82688
        .AddBezierSegment 20494, 82688, 19976, 83238, 19717, 83990, 19717, 84946
        .AddBezierSegment 19717, 84946, 19717, 85955, 19978, 86739, 20498, 87298
        .AddBezierSegment 20498, 87298, 21019, 87857, 21705, 88136, 22555, 88136
        .AddBezierSegment 22555, 88136, 23297, 88136, 23899, 87916, 24362, 87477
        .AddBezierSegment 24362, 87477, 24639, 87218, 24845, 86845, 24983, 86358
        .AddLineSegment 23772, 86070
        .AddBezierSegment 23772, 86070, 23700, 86384, 23551, 86632, 23323, 86814
        .AddBezierSegment 23323, 86814, 23095, 86996, 22818, 87087, 22494, 87087
        .AddBezierSegment 22494, 87087, 22044, 87087, 21678, 86926, 21399, 86602
        .AddBezierSegment 21399, 86602, 21118, 86281, 20978, 85759, 20978, 85036
        .AddBezierSegment 20978, 85036, 20978, 84271, 21117, 83725, 21392, 83401
        .AddBezierSegment 21392, 83401, 21668, 83075, 22027, 82912, 22468, 82912
        .AddBezierSegment 22468, 82912, 22794, 82912, 23075, 83015, 23309, 83223
        .AddBezierSegment 23309, 83223, 23544, 83430, 23713, 83755, 23815, 84199
        .CloseSubPath

        Set .brush = New waBrush
        .brush.SetColor RGB(43, 42, 41), 100
    End With
    MyCurves.Add path

    Set path = New waPath
    With path
        .StartSubPath
        .AddLineSegment -1245488, 81964
        .AddLineSegment -1245488, 88026
        .AddLineSegment -1243525, 88026
        .AddBezierSegment -1243525, 88026, -1242782, 88026, -1242298, 87997, -1242072, 87936
        .AddBezierSegment -1242072, 87936, -1241725, 87844, -1241434, 87647, -1241200, 87342
        .AddBezierSegment -1241200, 87342, -1240966, 87037, -1240848, 86643, -1240848, 86161
        .AddBezierSegment -1240848, 86161, -1240848, 85788, -1240916, 85476, -1241050, 85222
        .AddBezierSegment -1241050, 85222, -1241186, 84968, -1241356, 84769, -1241564, 84624
        .AddBezierSegment -1241564, 84624, -1241771, 84480, -1241983, 84384, -1242197, 84338
        .AddBezierSegment -1242197, 84338, -1242489, 84280, -1242911, 84250, -1243465, 84250
        .AddLineSegment -1244261, 84250
        .AddLineSegment -1244261, 81964
        .AddLineSegment -1245488, 81964
        .CloseSubPath

        .StartSubPath
        .AddLineSegment -1244261, 87002
        .AddLineSegment -1244261, 85275
        .AddLineSegment -1243593, 85275
        .AddBezierSegment -1243593, 85275, -1243112, 85275, -1242792, 85307, -1242629, 85371
        .AddBezierSegment -1242629, 85371, -1242467, 85434, -1242340, 85533, -1242248, 85670
        .AddBezierSegment -1242248, 85670, -1242157, 85805, -1242110, 85962, -1242110, 86142
        .AddBezierSegment -1242110, 86142, -1242110, 86364, -1242175, 86546, -1242303, 86690
        .AddBezierSegment -1242303, 86690, -1242433, 86834, -1242595, 86924, -1242794, 86961
        .AddBezierSegment -1242794, 86961, -1242940, 86988, -1243232, 87002, -1243672, 87002
        .AddLineSegment -1244261, 87002
        .CloseSubPath

        .StartSubPath
        .AddLineSegment -1239807, 81964
        .AddLineSegment -1239807, 87975
        .AddLineSegment -1238579, 87975
        .AddLineSegment -1238579, 82989
        .AddLineSegment -1235540, 82989
        .AddLineSegment -1235540, 81964
        .AddLineSegment -1239807, 81964
        .CloseSubPath

        .StartSubPath
        .AddLineSegment -1234923, 84959
        .AddBezierSegment -1234923, 84959, -1234923, 85577, -1234830, 86096, -1234645, 86516
        .AddBezierSegment -1234645, 86516, -1234508, 86825, -1234319, 87103, -1234081, 87349
        .AddBezierSegment -1234081, 87349, -1233842, 87594, -1233581, 87776, -1233297, 87896
        .AddBezierSegment -1233297, 87896, -1232919, 88056, -1232483, 88136, -1231989, 88136
        .AddBezierSegment -1231989, 88136, -1231095, 88136, -1230381, 87858, -1229843, 87304
        .AddBezierSegment -1229843, 87304, -1229307, 86748, -1229039, 85976, -1229039, 84987
        .AddBezierSegment -1229039, 84987, -1229039, 84006, -1229305, 83238, -1229838, 82685
        .AddBezierSegment -1229838, 82685, -1230370, 82131, -1231082, 81854, -1231973, 81854
        .AddBezierSegment -1231973, 81854, -1232874, 81854, -1233592, 82129, -1234124, 82681
        .AddBezierSegment -1234124, 82681, -1234656, 83231, -1234923, 83990, -1234923, 84959
        .CloseSubPath

        .StartSubPath
        .AddLineSegment -1233662, 84999
        .AddBezierSegment -1233662, 84999, -1233662, 84312, -1233502, 83791, -1233185, 83436
        .AddBezierSegment -1233185, 83436, -1232867, 83082, -1232464, 82904, -1231975, 82904
        .AddBezierSegment -1231975, 82904, -1231486, 82904, -1231085, 83080, -1230770, 83432
        .AddBezierSegment -1230770, 83432, -1230457, 83784, -1230300, 84312, -1230300, 85016
        .AddBezierSegment -1230300, 85016, -1230300, 85712, -1230453, 86231, -1230759, 86573
        .AddBezierSegment -1230759, 86573, -1231064, 86916, -1231469, 87087, -1231975, 87087
        .AddBezierSegment -1231975, 87087, -1232480, 87087, -1232888, 86913, -1233197, 86567
        .AddBezierSegment -1233197, 86567, -1233506, 86220, -1233662, 85698, -1233662, 84999
        .CloseSubPath

        .StartSubPath
        .AddLineSegment -1226720, 81964
        .AddLineSegment -1226720, 87002
        .AddLineSegment -1228515, 87002
        .AddLineSegment -1228515, 88026
        .AddLineSegment -1223706, 88026
        .AddLineSegment -1223706, 87002
        .AddLineSegment -1225493, 87002
        .AddLineSegment -1225493, 81964
        .AddLineSegment -1226720, 81964
        .CloseSubPath

        .StartSubPath
        .AddLineSegment -1219034, 84199
        .AddLineSegment -1217849, 83835
        .AddBezierSegment -1217849, 83835, -1218031, 83168, -1218334, 82672, -1218757, 82348
        .AddBezierSegment -1218757, 82348, -1219181, 82025, -1219717, 81862, -1220368, 81862
        .AddBezierSegment -1220368, 81862, -1221173, 81862, -1221835, 82138, -1222354, 82688
        .AddBezierSegment -1222354, 82688, -1222872, 83238, -1223132, 83990, -1223132, 84946
        .AddBezierSegment -1223132, 84946, -1223132, 85955, -1222871, 86739, -1222350, 87298
        .AddBezierSegment -1222350, 87298, -1221829, 87857, -1221144, 88136, -1220294, 88136
        .AddBezierSegment -1220294, 88136, -1219552, 88136, -1218949, 87916, -1218486, 87477
        .AddBezierSegment -1218486, 87477, -1218210, 87218, -1218004, 86845, -1217866, 86358
        .AddLineSegment -1219076, 86070
        .AddBezierSegment -1219076, 86070, -1219148, 86384, -1219298, 86632, -1219525, 86814
        .AddBezierSegment -1219525, 86814, -1219754, 86996, -1220030, 87087, -1220355, 87087
        .AddBezierSegment -1220355, 87087, -1220805, 87087, -1221170, 86926, -1221450, 86602
        .AddBezierSegment -1221450, 86602, -1221731, 86281, -1221870, 85759, -1221870, 85036
        .AddBezierSegment -1221870, 85036, -1221870, 84271, -1221732, 83725, -1221457, 83401
        .AddBezierSegment -1221457, 83401, -1221180, 83075, -1220822, 82912, -1220380, 82912
        .AddBezierSegment -1220380, 82912, -1220054, 82912, -1219773, 83015, -1219539, 83223
        .AddBezierSegment -1219539, 83223, -1219305, 83430, -1219136, 83755, -1219034, 84199
        .CloseSubPath

        .StartSubPath
        .AddLineSegment -1211336, 81964
        .AddLineSegment -1212659, 81964
        .AddLineSegment -1213182, 83344
        .AddLineSegment -1215613, 83344
        .AddLineSegment -1216117, 81964
        .AddLineSegment -1217415, 81964
        .AddLineSegment -1215062, 88026
        .AddLineSegment -1213765, 88026
        .AddLineSegment -1211336, 81964
        .CloseSubPath

        .StartSubPath
        .AddLineSegment -1213580, 84369
        .AddLineSegment -1214422, 86612
        .AddLineSegment -1215241, 84369
        .AddLineSegment -1213580, 84369
        .CloseSubPath

        .StartSubPath
        .AddLineSegment -1210649, 81964
        .AddLineSegment -1210649, 87975
        .AddLineSegment -1209421, 87975
        .AddLineSegment -1209421, 82989
        .AddLineSegment -1206382, 82989
        .AddLineSegment -1206382, 81964
        .AddLineSegment -1210649, 81964
        .CloseSubPath

        .StartSubPath
        .AddLineSegment -1201633, 84199
        .AddLineSegment -1200448, 83835
        .AddBezierSegment -1200448, 83835, -1200630, 83168, -1200934, 82672, -1201357, 82348
        .AddBezierSegment -1201357, 82348, -1201780, 82025, -1202316, 81862, -1202967, 81862
        .AddBezierSegment -1202967, 81862, -1203773, 81862, -1204435, 82138, -1204954, 82688
        .AddBezierSegment -1204954, 82688, -1205472, 83238, -1205731, 83990, -1205731, 84946
        .AddBezierSegment -1205731, 84946, -1205731, 85955, -1205470, 86739, -1204950, 87298
        .AddBezierSegment -1204950, 87298, -1204429, 87857, -1203743, 88136, -1202894, 88136
        .AddBezierSegment -1202894, 88136, -1202151, 88136, -1201549, 87916, -1201086, 87477
        .AddBezierSegment -1201086, 87477, -1200809, 87218, -1200603, 86845, -1200465, 86358
        .AddLineSegment -1201676, 86070
        .AddBezierSegment -1201676, 86070, -1201748, 86384, -1201897, 86632, -1202125, 86814
        .AddBezierSegment -1202125, 86814, -1202353, 86996, -1202630, 87087, -1202954, 87087
        .AddBezierSegment -1202954, 87087, -1203404, 87087, -1203770, 86926, -1204049, 86602
        .AddBezierSegment -1204049, 86602, -1204330, 86281, -1204470, 85759, -1204470, 85036
        .AddBezierSegment -1204470, 85036, -1204470, 84271, -1204332, 83725, -1204056, 83401
        .AddBezierSegment -1204056, 83401, -1203780, 83075, -1203421, 82912, -1202980, 82912
        .AddBezierSegment -1202980, 82912, -1202654, 82912, -1202373, 83015, -1202139, 83223
        .AddBezierSegment -1202139, 83223, -1201904, 83430, -1201735, 83755, -1201633, 84199
        .CloseSubPath

        Set .brush = New waBrush
        .brush.SetColor RGB(43, 42, 41), 100
    End With
    MyCurves.Add path

    Set path = New waPath
    With path
        .StartSubPath
        .AddLineSegment -1049202, 81964
        .AddLineSegment -1049202, 88026
        .AddLineSegment -1047239, 88026
        .AddBezierSegment -1047239, 88026, -1046495, 88026, -1046011, 87997, -1045785, 87936
        .AddBezierSegment -1045785, 87936, -1045438, 87844, -1045148, 87647, -1044913, 87342
        .AddBezierSegment -1044913, 87342, -1044679, 87037, -1044562, 86643, -1044562, 86161
        .AddBezierSegment -1044562, 86161, -1044562, 85788, -1044630, 85476, -1044764, 85222
        .AddBezierSegment -1044764, 85222, -1044899, 84968, -1045070, 84769, -1045277, 84624
        .AddBezierSegment -1045277, 84624, -1045485, 84480, -1045697, 84384, -1045911, 84338
        .AddBezierSegment -1045911, 84338, -1046203, 84280, -1046625, 84250, -1047178, 84250
        .AddLineSegment -1047974, 84250
        .AddLineSegment -1047974, 81964
        .AddLineSegment -1049202, 81964
        .CloseSubPath

        .StartSubPath
        .AddLineSegment -1047974, 87002
        .AddLineSegment -1047974, 85275
        .AddLineSegment -1047307, 85275
        .AddBezierSegment -1047307, 85275, -1046825, 85275, -1046505, 85307, -1046343, 85371
        .AddBezierSegment -1046343, 85371, -1046181, 85434, -1046054, 85533, -1045962, 85670
        .AddBezierSegment -1045962, 85670, -1045870, 85805, -1045824, 85962, -1045824, 86142
        .AddBezierSegment -1045824, 86142, -1045824, 86364, -1045889, 86546, -1046017, 86690
        .AddBezierSegment -1046017, 86690, -1046147, 86834, -1046309, 86924, -1046508, 86961
        .AddBezierSegment -1046508, 86961, -1046653, 86988, -1046946, 87002, -1047386, 87002
        .AddLineSegment -1047974, 87002
        .CloseSubPath

        .StartSubPath
        .AddLineSegment -1043521, 81964
        .AddLineSegment -1043521, 87975
        .AddLineSegment -1042293, 87975
        .AddLineSegment -1042293, 82989
        .AddLineSegment -1039253, 82989
        .AddLineSegment -1039253, 81964
        .AddLineSegment -1043521, 81964
        .CloseSubPath

        .StartSubPath
        .AddLineSegment -1038637, 84959
        .AddBezierSegment -1038637, 84959, -1038637, 85577, -1038544, 86096, -1038359, 86516
        .AddBezierSegment -1038359, 86516, -1038222, 86825, -1038033, 87103, -1037794, 87349
        .AddBezierSegment -1037794, 87349, -1037556, 87594, -1037295, 87776, -1037011, 87896
        .AddBezierSegment -1037011, 87896, -1036633, 88056, -1036197, 88136, -1035703, 88136
        .AddBezierSegment -1035703, 88136, -1034809, 88136, -1034094, 87858, -1033557, 87304
        .AddBezierSegment -1033557, 87304, -1033021, 86748, -1032753, 85976, -1032753, 84987
        .AddBezierSegment -1032753, 84987, -1032753, 84006, -1033019, 83238, -1033551, 82685
        .AddBezierSegment -1033551, 82685, -1034083, 82131, -1034796, 81854, -1035686, 81854
        .AddBezierSegment -1035686, 81854, -1036588, 81854, -1037306, 82129, -1037838, 82681
        .AddBezierSegment -1037838, 82681, -1038370, 83231, -1038637, 83990, -1038637, 84959
        .CloseSubPath

        .StartSubPath
        .AddLineSegment -1037375, 84999
        .AddBezierSegment -1037375, 84999, -1037375, 84312, -1037216, 83791, -1036898, 83436
        .AddBezierSegment -1036898, 83436, -1036581, 83082, -1036177, 82904, -1035689, 82904
        .AddBezierSegment -1035689, 82904, -1035199, 82904, -1034799, 83080, -1034484, 83432
        .AddBezierSegment -1034484, 83432, -1034171, 83784, -1034014, 84312, -1034014, 85016
        .AddBezierSegment -1034014, 85016, -1034014, 85712, -1034166, 86231, -1034473, 86573
        .AddBezierSegment -1034473, 86573, -1034777, 86916, -1035182, 87087, -1035689, 87087
        .AddBezierSegment -1035689, 87087, -1036194, 87087, -1036602, 86913, -1036911, 86567
        .AddBezierSegment -1036911, 86567, -1037220, 86220, -1037375, 85698, -1037375, 84999
        .CloseSubPath

        .StartSubPath
        .AddLineSegment -1030434, 81964
        .AddLineSegment -1030434, 87002
        .AddLineSegment -1032229, 87002
        .AddLineSegment -1032229, 88026
        .AddLineSegment -1027420, 88026
        .AddLineSegment -1027420, 87002
        .AddLineSegment -1029206, 87002
        .AddLineSegment -1029206, 81964
        .AddLineSegment -1030434, 81964
        .CloseSubPath

        .StartSubPath
        .AddLineSegment -1022748, 84199
        .AddLineSegment -1021562, 83835
        .AddBezierSegment -1021562, 83835, -1021744, 83168, -1022048, 82672, -1022471, 82348
        .AddBezierSegment -1022471, 82348, -1022894, 82025, -1023431, 81862, -1024081, 81862
        .AddBezierSegment -1024081, 81862, -1024887, 81862, -1025549, 82138, -1026068, 82688
        .AddBezierSegment -1026068, 82688, -1026586, 83238, -1026846, 83990, -1026846, 84946
        .AddBezierSegment -1026846, 84946, -1026846, 85955, -1026584, 86739, -1026064, 87298
        .AddBezierSegment -1026064, 87298, -1025543, 87857, -1024857, 88136, -1024008, 88136
        .AddBezierSegment -1024008, 88136, -1023266, 88136, -1022663, 87916, -1022200, 87477
        .AddBezierSegment -1022200, 87477, -1021924, 87218, -1021718, 86845, -1021579, 86358
        .AddLineSegment -1022790, 86070
        .AddBezierSegment -1022790, 86070, -1022862, 86384, -1023012, 86632, -1023239, 86814
        .AddBezierSegment -1023239, 86814, -1023467, 86996, -1023744, 87087, -1024069, 87087
        .AddBezierSegment -1024069, 87087, -1024519, 87087, -1024884, 86926, -1025164, 86602
        .AddBezierSegment -1025164, 86602, -1025444, 86281, -1025584, 85759, -1025584, 85036
        .AddBezierSegment -1025584, 85036, -1025584, 84271, -1025446, 83725, -1025171, 83401
        .AddBezierSegment -1025171, 83401, -1024894, 83075, -1024536, 82912, -1024094, 82912
        .AddBezierSegment -1024094, 82912, -1023768, 82912, -1023487, 83015, -1023253, 83223
        .AddBezierSegment -1023253, 83223, -1023019, 83430, -1022849, 83755, -1022748, 84199
        .CloseSubPath

        .StartSubPath
        .AddLineSegment -1015050, 81964
        .AddLineSegment -1016372, 81964
        .AddLineSegment -1016896, 83344
        .AddLineSegment -1019327, 83344
        .AddLineSegment -1019831, 81964
        .AddLineSegment -1021129, 81964
        .AddLineSegment -1018775, 88026
        .AddLineSegment -1017479, 88026
        .AddLineSegment -1015050, 81964
        .CloseSubPath

        .StartSubPath
        .AddLineSegment -1017294, 84369
        .AddLineSegment -1018136, 86612
        .AddLineSegment -1018955, 84369
        .AddLineSegment -1017294, 84369
        .CloseSubPath

        .StartSubPath
        .AddLineSegment -1014363, 81964
        .AddLineSegment -1014363, 87975
        .AddLineSegment -1013135, 87975
        .AddLineSegment -1013135, 82989
        .AddLineSegment -1010096, 82989
        .AddLineSegment -1010096, 81964
        .AddLineSegment -1014363, 81964
        .CloseSubPath

        .StartSubPath
        .AddLineSegment -1005347, 84199
        .AddLineSegment -1004162, 83835
        .AddBezierSegment -1004162, 83835, -1004344, 83168, -1004647, 82672, -1005071, 82348
        .AddBezierSegment -1005071, 82348, -1005494, 82025, -1006030, 81862, -1006681, 81862
        .AddBezierSegment -1006681, 81862, -1007486, 81862, -1008148, 82138, -1008668, 82688
        .AddBezierSegment -1008668, 82688, -1009185, 83238, -1009445, 83990, -1009445, 84946
        .AddBezierSegment -1009445, 84946, -1009445, 85955, -1009184, 86739, -1008663, 87298
        .AddBezierSegment -1008663, 87298, -1008143, 87857, -1007457, 88136, -1006607, 88136
        .AddBezierSegment -1006607, 88136, -1005865, 88136, -1005263, 87916, -1004800, 87477
        .AddBezierSegment -1004800, 87477, -1004523, 87218, -1004317, 86845, -1004179, 86358
        .AddLineSegment -1005389, 86070
        .AddBezierSegment -1005389, 86070, -1005462, 86384, -1005611, 86632, -1005838, 86814
        .AddBezierSegment -1005838, 86814, -1006067, 86996, -1006343, 87087, -1006668, 87087
        .AddBezierSegment -1006668, 87087, -1007118, 87087, -1007484, 86926, -1007763, 86602
        .AddBezierSegment -1007763, 86602, -1008044, 86281, -1008183, 85759, -1008183, 85036
        .AddBezierSegment -1008183, 85036, -1008183, 84271, -1008045, 83725, -1007770, 83401
        .AddBezierSegment -1007770, 83401, -1007494, 83075, -1007135, 82912, -1006693, 82912
        .AddBezierSegment -1006693, 82912, -1006367, 82912, -1006087, 83015, -1005852, 83223
        .AddBezierSegment -1005852, 83223, -1005618, 83430, -1005449, 83755, -1005347, 84199
        .CloseSubPath

        Set .brush = New waBrush
        .brush.SetColor RGB(43, 42, 41), 100
    End With
    MyCurves.Add path

    Set path = New waPath
    With path
        .StartSubPath
        .AddLineSegment -1457403, 81964
        .AddLineSegment -1457403, 88026
        .AddLineSegment -1455441, 88026
        .AddBezierSegment -1455441, 88026, -1454697, 88026, -1454213, 87997, -1453987, 87936
        .AddBezierSegment -1453987, 87936, -1453640, 87844, -1453349, 87647, -1453115, 87342
        .AddBezierSegment -1453115, 87342, -1452881, 87037, -1452764, 86643, -1452764, 86161
        .AddBezierSegment -1452764, 86161, -1452764, 85788, -1452832, 85476, -1452966, 85222
        .AddBezierSegment -1452966, 85222, -1453101, 84968, -1453272, 84769, -1453479, 84624
        .AddBezierSegment -1453479, 84624, -1453687, 84480, -1453898, 84384, -1454113, 84338
        .AddBezierSegment -1454113, 84338, -1454405, 84280, -1454827, 84250, -1455380, 84250
        .AddLineSegment -1456176, 84250
        .AddLineSegment -1456176, 81964
        .AddLineSegment -1457403, 81964
        .CloseSubPath

        .StartSubPath
        .AddLineSegment -1456176, 87002
        .AddLineSegment -1456176, 85275
        .AddLineSegment -1455508, 85275
        .AddBezierSegment -1455508, 85275, -1455027, 85275, -1454707, 85307, -1454545, 85371
        .AddBezierSegment -1454545, 85371, -1454382, 85434, -1454255, 85533, -1454164, 85670
        .AddBezierSegment -1454164, 85670, -1454072, 85805, -1454025, 85962, -1454025, 86142
        .AddBezierSegment -1454025, 86142, -1454025, 86364, -1454090, 86546, -1454219, 86690
        .AddBezierSegment -1454219, 86690, -1454349, 86834, -1454511, 86924, -1454710, 86961
        .AddBezierSegment -1454710, 86961, -1454855, 86988, -1455147, 87002, -1455587, 87002
        .AddLineSegment -1456176, 87002
        .CloseSubPath

        .StartSubPath
        .AddLineSegment -1451722, 81964
        .AddLineSegment -1451722, 87975
        .AddLineSegment -1450495, 87975
        .AddLineSegment -1450495, 82989
        .AddLineSegment -1447455, 82989
        .AddLineSegment -1447455, 81964
        .AddLineSegment -1451722, 81964
        .CloseSubPath

        .StartSubPath
        .AddLineSegment -1446839, 84959
        .AddBezierSegment -1446839, 84959, -1446839, 85577, -1446745, 86096, -1446561, 86516
        .AddBezierSegment -1446561, 86516, -1446424, 86825, -1446235, 87103, -1445996, 87349
        .AddBezierSegment -1445996, 87349, -1445758, 87594, -1445497, 87776, -1445213, 87896
        .AddBezierSegment -1445213, 87896, -1444835, 88056, -1444399, 88136, -1443905, 88136
        .AddBezierSegment -1443905, 88136, -1443010, 88136, -1442296, 87858, -1441758, 87304
        .AddBezierSegment -1441758, 87304, -1441222, 86748, -1440954, 85976, -1440954, 84987
        .AddBezierSegment -1440954, 84987, -1440954, 84006, -1441221, 83238, -1441753, 82685
        .AddBezierSegment -1441753, 82685, -1442285, 82131, -1442997, 81854, -1443888, 81854
        .AddBezierSegment -1443888, 81854, -1444790, 81854, -1445508, 82129, -1446040, 82681
        .AddBezierSegment -1446040, 82681, -1446572, 83231, -1446839, 83990, -1446839, 84959
        .CloseSubPath

        .StartSubPath
        .AddLineSegment -1445577, 84999
        .AddBezierSegment -1445577, 84999, -1445577, 84312, -1445418, 83791, -1445100, 83436
        .AddBezierSegment -1445100, 83436, -1444783, 83082, -1444379, 82904, -1443891, 82904
        .AddBezierSegment -1443891, 82904, -1443401, 82904, -1443000, 83080, -1442686, 83432
        .AddBezierSegment -1442686, 83432, -1442372, 83784, -1442216, 84312, -1442216, 85016
        .AddBezierSegment -1442216, 85016, -1442216, 85712, -1442368, 86231, -1442674, 86573
        .AddBezierSegment -1442674, 86573, -1442979, 86916, -1443384, 87087, -1443891, 87087
        .AddBezierSegment -1443891, 87087, -1444396, 87087, -1444804, 86913, -1445113, 86567
        .AddBezierSegment -1445113, 86567, -1445422, 86220, -1445577, 85698, -1445577, 84999
        .CloseSubPath

        .StartSubPath
        .AddLineSegment -1438636, 81964
        .AddLineSegment -1438636, 87002
        .AddLineSegment -1440431, 87002
        .AddLineSegment -1440431, 88026
        .AddLineSegment -1435621, 88026
        .AddLineSegment -1435621, 87002
        .AddLineSegment -1437408, 87002
        .AddLineSegment -1437408, 81964
        .AddLineSegment -1438636, 81964
        .CloseSubPath

        .StartSubPath
        .AddLineSegment -1430949, 84199
        .AddLineSegment -1429764, 83835
        .AddBezierSegment -1429764, 83835, -1429946, 83168, -1430249, 82672, -1430673, 82348
        .AddBezierSegment -1430673, 82348, -1431096, 82025, -1431632, 81862, -1432283, 81862
        .AddBezierSegment -1432283, 81862, -1433089, 81862, -1433750, 82138, -1434270, 82688
        .AddBezierSegment -1434270, 82688, -1434788, 83238, -1435047, 83990, -1435047, 84946
        .AddBezierSegment -1435047, 84946, -1435047, 85955, -1434786, 86739, -1434265, 87298
        .AddBezierSegment -1434265, 87298, -1433745, 87857, -1433059, 88136, -1432209, 88136
        .AddBezierSegment -1432209, 88136, -1431467, 88136, -1430865, 87916, -1430402, 87477
        .AddBezierSegment -1430402, 87477, -1430125, 87218, -1429919, 86845, -1429781, 86358
        .AddLineSegment -1430992, 86070
        .AddBezierSegment -1430992, 86070, -1431064, 86384, -1431213, 86632, -1431440, 86814
        .AddBezierSegment -1431440, 86814, -1431669, 86996, -1431946, 87087, -1432270, 87087
        .AddBezierSegment -1432270, 87087, -1432720, 87087, -1433086, 86926, -1433365, 86602
        .AddBezierSegment -1433365, 86602, -1433646, 86281, -1433786, 85759, -1433786, 85036
        .AddBezierSegment -1433786, 85036, -1433786, 84271, -1433647, 83725, -1433372, 83401
        .AddBezierSegment -1433372, 83401, -1433096, 83075, -1432737, 82912, -1432296, 82912
        .AddBezierSegment -1432296, 82912, -1431970, 82912, -1431689, 83015, -1431455, 83223
        .AddBezierSegment -1431455, 83223, -1431220, 83430, -1431051, 83755, -1430949, 84199
        .CloseSubPath

        .StartSubPath
        .AddLineSegment -1423252, 81964
        .AddLineSegment -1424574, 81964
        .AddLineSegment -1425097, 83344
        .AddLineSegment -1427529, 83344
        .AddLineSegment -1428033, 81964
        .AddLineSegment -1429331, 81964
        .AddLineSegment -1426977, 88026
        .AddLineSegment -1425680, 88026
        .AddLineSegment -1423252, 81964
        .CloseSubPath

        .StartSubPath
        .AddLineSegment -1425495, 84369
        .AddLineSegment -1426338, 86612
        .AddLineSegment -1427156, 84369
        .AddLineSegment -1425495, 84369
        .CloseSubPath

        .StartSubPath
        .AddLineSegment -1422564, 81964
        .AddLineSegment -1422564, 87975
        .AddLineSegment -1421337, 87975
        .AddLineSegment -1421337, 82989
        .AddLineSegment -1418297, 82989
        .AddLineSegment -1418297, 81964
        .AddLineSegment -1422564, 81964
        .CloseSubPath

        .StartSubPath
        .AddLineSegment -1413549, 84199
        .AddLineSegment -1412363, 83835
        .AddBezierSegment -1412363, 83835, -1412546, 83168, -1412849, 82672, -1413272, 82348
        .AddBezierSegment -1413272, 82348, -1413696, 82025, -1414232, 81862, -1414882, 81862
        .AddBezierSegment -1414882, 81862, -1415688, 81862, -1416350, 82138, -1416869, 82688
        .AddBezierSegment -1416869, 82688, -1417387, 83238, -1417647, 83990, -1417647, 84946
        .AddBezierSegment -1417647, 84946, -1417647, 85955, -1417386, 86739, -1416865, 87298
        .AddBezierSegment -1416865, 87298, -1416344, 87857, -1415659, 88136, -1414809, 88136
        .AddBezierSegment -1414809, 88136, -1414067, 88136, -1413464, 87916, -1413001, 87477
        .AddBezierSegment -1413001, 87477, -1412725, 87218, -1412519, 86845, -1412380, 86358
        .AddLineSegment -1413591, 86070
        .AddBezierSegment -1413591, 86070, -1413663, 86384, -1413813, 86632, -1414040, 86814
        .AddBezierSegment -1414040, 86814, -1414269, 86996, -1414545, 87087, -1414870, 87087
        .AddBezierSegment -1414870, 87087, -1415320, 87087, -1415685, 86926, -1415965, 86602
        .AddBezierSegment -1415965, 86602, -1416245, 86281, -1416385, 85759, -1416385, 85036
        .AddBezierSegment -1416385, 85036, -1416385, 84271, -1416247, 83725, -1415972, 83401
        .AddBezierSegment -1415972, 83401, -1415695, 83075, -1415337, 82912, -1414895, 82912
        .AddBezierSegment -1414895, 82912, -1414569, 82912, -1414288, 83015, -1414054, 83223
        .AddBezierSegment -1414054, 83223, -1413820, 83430, -1413650, 83755, -1413549, 84199
        .CloseSubPath

        Set .brush = New waBrush
        .brush.SetColor RGB(43, 42, 41), 100
    End With
    MyCurves.Add path

    Set path = New waPath
    With path
        .StartSubPath
        .AddLineSegment -402048, 81964
        .AddLineSegment -402048, 88026
        .AddLineSegment -400085, 88026
        .AddBezierSegment -400085, 88026, -399341, 88026, -398857, 87997, -398632, 87936
        .AddBezierSegment -398632, 87936, -398284, 87844, -397994, 87647, -397759, 87342
        .AddBezierSegment -397759, 87342, -397525, 87037, -397408, 86643, -397408, 86161
        .AddBezierSegment -397408, 86161, -397408, 85788, -397476, 85476, -397610, 85222
        .AddBezierSegment -397610, 85222, -397745, 84968, -397916, 84769, -398123, 84624
        .AddBezierSegment -398123, 84624, -398331, 84480, -398543, 84384, -398757, 84338
        .AddBezierSegment -398757, 84338, -399049, 84280, -399471, 84250, -400024, 84250
        .AddLineSegment -400820, 84250
        .AddLineSegment -400820, 81964
        .AddLineSegment -402048, 81964
        .CloseSubPath

        .StartSubPath
        .AddLineSegment -400820, 87002
        .AddLineSegment -400820, 85275
        .AddLineSegment -400153, 85275
        .AddBezierSegment -400153, 85275, -399671, 85275, -399351, 85307, -399189, 85371
        .AddBezierSegment -399189, 85371, -399027, 85434, -398900, 85533, -398808, 85670
        .AddBezierSegment -398808, 85670, -398716, 85805, -398670, 85962, -398670, 86142
        .AddBezierSegment -398670, 86142, -398670, 86364, -398735, 86546, -398863, 86690
        .AddBezierSegment -398863, 86690, -398993, 86834, -399155, 86924, -399354, 86961
        .AddBezierSegment -399354, 86961, -399499, 86988, -399791, 87002, -400232, 87002
        .AddLineSegment -400820, 87002
        .CloseSubPath

        .StartSubPath
        .AddLineSegment -396367, 81964
        .AddLineSegment -396367, 87975
        .AddLineSegment -395139, 87975
        .AddLineSegment -395139, 82989
        .AddLineSegment -392099, 82989
        .AddLineSegment -392099, 81964
        .AddLineSegment -396367, 81964
        .CloseSubPath

        .StartSubPath
        .AddLineSegment -391483, 84959
        .AddBezierSegment -391483, 84959, -391483, 85577, -391390, 86096, -391205, 86516
        .AddBezierSegment -391205, 86516, -391068, 86825, -390879, 87103, -390640, 87349
        .AddBezierSegment -390640, 87349, -390402, 87594, -390141, 87776, -389857, 87896
        .AddBezierSegment -389857, 87896, -389479, 88056, -389043, 88136, -388549, 88136
        .AddBezierSegment -388549, 88136, -387654, 88136, -386940, 87858, -386403, 87304
        .AddBezierSegment -386403, 87304, -385867, 86748, -385599, 85976, -385599, 84987
        .AddBezierSegment -385599, 84987, -385599, 84006, -385865, 83238, -386397, 82685
        .AddBezierSegment -386397, 82685, -386929, 82131, -387642, 81854, -388532, 81854
        .AddBezierSegment -388532, 81854, -389434, 81854, -390152, 82129, -390684, 82681
        .AddBezierSegment -390684, 82681, -391216, 83231, -391483, 83990, -391483, 84959
        .CloseSubPath

        .StartSubPath
        .AddLineSegment -390221, 84999
        .AddBezierSegment -390221, 84999, -390221, 84312, -390062, 83791, -389744, 83436
        .AddBezierSegment -389744, 83436, -389427, 83082, -389023, 82904, -388535, 82904
        .AddBezierSegment -388535, 82904, -388045, 82904, -387645, 83080, -387330, 83432
        .AddBezierSegment -387330, 83432, -387017, 83784, -386860, 84312, -386860, 85016
        .AddBezierSegment -386860, 85016, -386860, 85712, -387012, 86231, -387319, 86573
        .AddBezierSegment -387319, 86573, -387623, 86916, -388028, 87087, -388535, 87087
        .AddBezierSegment -388535, 87087, -389040, 87087, -389448, 86913, -389757, 86567
        .AddBezierSegment -389757, 86567, -390066, 86220, -390221, 85698, -390221, 84999
        .CloseSubPath

        .StartSubPath
        .AddLineSegment -383280, 81964
        .AddLineSegment -383280, 87002
        .AddLineSegment -385075, 87002
        .AddLineSegment -385075, 88026
        .AddLineSegment -380266, 88026
        .AddLineSegment -380266, 87002
        .AddLineSegment -382052, 87002
        .AddLineSegment -382052, 81964
        .AddLineSegment -383280, 81964
        .CloseSubPath

        .StartSubPath
        .AddLineSegment -375594, 84199
        .AddLineSegment -374408, 83835
        .AddBezierSegment -374408, 83835, -374590, 83168, -374894, 82672, -375317, 82348
        .AddBezierSegment -375317, 82348, -375740, 82025, -376277, 81862, -376927, 81862
        .AddBezierSegment -376927, 81862, -377733, 81862, -378395, 82138, -378914, 82688
        .AddBezierSegment -378914, 82688, -379432, 83238, -379692, 83990, -379692, 84946
        .AddBezierSegment -379692, 84946, -379692, 85955, -379430, 86739, -378910, 87298
        .AddBezierSegment -378910, 87298, -378389, 87857, -377703, 88136, -376854, 88136
        .AddBezierSegment -376854, 88136, -376112, 88136, -375509, 87916, -375046, 87477
        .AddBezierSegment -375046, 87477, -374770, 87218, -374564, 86845, -374425, 86358
        .AddLineSegment -375636, 86070
        .AddBezierSegment -375636, 86070, -375708, 86384, -375858, 86632, -376085, 86814
        .AddBezierSegment -376085, 86814, -376313, 86996, -376590, 87087, -376914, 87087
        .AddBezierSegment -376914, 87087, -377365, 87087, -377730, 86926, -378010, 86602
        .AddBezierSegment -378010, 86602, -378290, 86281, -378430, 85759, -378430, 85036
        .AddBezierSegment -378430, 85036, -378430, 84271, -378292, 83725, -378017, 83401
        .AddBezierSegment -378017, 83401, -377740, 83075, -377382, 82912, -376940, 82912
        .AddBezierSegment -376940, 82912, -376614, 82912, -376333, 83015, -376099, 83223
        .AddBezierSegment -376099, 83223, -375865, 83430, -375695, 83755, -375594, 84199
        .CloseSubPath

        .StartSubPath
        .AddLineSegment -367896, 81964
        .AddLineSegment -369218, 81964
        .AddLineSegment -369742, 83344
        .AddLineSegment -372173, 83344
        .AddLineSegment -372677, 81964
        .AddLineSegment -373975, 81964
        .AddLineSegment -371621, 88026
        .AddLineSegment -370325, 88026
        .AddLineSegment -367896, 81964
        .CloseSubPath

        .StartSubPath
        .AddLineSegment -370140, 84369
        .AddLineSegment -370982, 86612
        .AddLineSegment -371801, 84369
        .AddLineSegment -370140, 84369
        .CloseSubPath

        .StartSubPath
        .AddLineSegment -367209, 81964
        .AddLineSegment -367209, 87975
        .AddLineSegment -365981, 87975
        .AddLineSegment -365981, 82989
        .AddLineSegment -362942, 82989
        .AddLineSegment -362942, 81964
        .AddLineSegment -367209, 81964
        .CloseSubPath

        .StartSubPath
        .AddLineSegment -358193, 84199
        .AddLineSegment -357008, 83835
        .AddBezierSegment -357008, 83835, -357190, 83168, -357493, 82672, -357917, 82348
        .AddBezierSegment -357917, 82348, -358340, 82025, -358876, 81862, -359527, 81862
        .AddBezierSegment -359527, 81862, -360332, 81862, -360994, 82138, -361514, 82688
        .AddBezierSegment -361514, 82688, -362031, 83238, -362291, 83990, -362291, 84946
        .AddBezierSegment -362291, 84946, -362291, 85955, -362030, 86739, -361509, 87298
        .AddBezierSegment -361509, 87298, -360989, 87857, -360303, 88136, -359453, 88136
        .AddBezierSegment -359453, 88136, -358711, 88136, -358108, 87916, -357646, 87477
        .AddBezierSegment -357646, 87477, -357369, 87218, -357163, 86845, -357025, 86358
        .AddLineSegment -358236, 86070
        .AddBezierSegment -358236, 86070, -358308, 86384, -358457, 86632, -358684, 86814
        .AddBezierSegment -358684, 86814, -358913, 86996, -359189, 87087, -359514, 87087
        .AddBezierSegment -359514, 87087, -359964, 87087, -360330, 86926, -360609, 86602
        .AddBezierSegment -360609, 86602, -360890, 86281, -361029, 85759, -361029, 85036
        .AddBezierSegment -361029, 85036, -361029, 84271, -360891, 83725, -360616, 83401
        .AddBezierSegment -360616, 83401, -360339, 83075, -359981, 82912, -359539, 82912
        .AddBezierSegment -359539, 82912, -359213, 82912, -358933, 83015, -358698, 83223
        .AddBezierSegment -358698, 83223, -358464, 83430, -358295, 83755, -358193, 84199
        .CloseSubPath

        Set .brush = New waBrush
        .brush.SetColor RGB(43, 42, 41), 100
    End With
    MyCurves.Add path

    Set path = New waPath
    With path
        .StartSubPath
        .AddLineSegment -205762, 81964
        .AddLineSegment -205762, 88026
        .AddLineSegment -203799, 88026
        .AddBezierSegment -203799, 88026, -203055, 88026, -202571, 87997, -202345, 87936
        .AddBezierSegment -202345, 87936, -201998, 87844, -201707, 87647, -201473, 87342
        .AddBezierSegment -201473, 87342, -201239, 87037, -201122, 86643, -201122, 86161
        .AddBezierSegment -201122, 86161, -201122, 85788, -201190, 85476, -201324, 85222
        .AddBezierSegment -201324, 85222, -201459, 84968, -201630, 84769, -201837, 84624
        .AddBezierSegment -201837, 84624, -202045, 84480, -202256, 84384, -202471, 84338
        .AddBezierSegment -202471, 84338, -202763, 84280, -203185, 84250, -203738, 84250
        .AddLineSegment -204534, 84250
        .AddLineSegment -204534, 81964
        .AddLineSegment -205762, 81964
        .CloseSubPath

        .StartSubPath
        .AddLineSegment -204534, 87002
        .AddLineSegment -204534, 85275
        .AddLineSegment -203866, 85275
        .AddBezierSegment -203866, 85275, -203385, 85275, -203065, 85307, -202903, 85371
        .AddBezierSegment -202903, 85371, -202740, 85434, -202613, 85533, -202522, 85670
        .AddBezierSegment -202522, 85670, -202430, 85805, -202383, 85962, -202383, 86142
        .AddBezierSegment -202383, 86142, -202383, 86364, -202448, 86546, -202577, 86690
        .AddBezierSegment -202577, 86690, -202707, 86834, -202869, 86924, -203068, 86961
        .AddBezierSegment -203068, 86961, -203213, 86988, -203505, 87002, -203945, 87002
        .AddLineSegment -204534, 87002
        .CloseSubPath

        .StartSubPath
        .AddLineSegment -200080, 81964
        .AddLineSegment -200080, 87975
        .AddLineSegment -198853, 87975
        .AddLineSegment -198853, 82989
        .AddLineSegment -195813, 82989
        .AddLineSegment -195813, 81964
        .AddLineSegment -200080, 81964
        .CloseSubPath

        .StartSubPath
        .AddLineSegment -195196, 84959
        .AddBezierSegment -195196, 84959, -195196, 85577, -195103, 86096, -194918, 86516
        .AddBezierSegment -194918, 86516, -194782, 86825, -194593, 87103, -194354, 87349
        .AddBezierSegment -194354, 87349, -194116, 87594, -193855, 87776, -193571, 87896
        .AddBezierSegment -193571, 87896, -193193, 88056, -192757, 88136, -192263, 88136
        .AddBezierSegment -192263, 88136, -191368, 88136, -190654, 87858, -190117, 87304
        .AddBezierSegment -190117, 87304, -189580, 86748, -189312, 85976, -189312, 84987
        .AddBezierSegment -189312, 84987, -189312, 84006, -189579, 83238, -190111, 82685
        .AddBezierSegment -190111, 82685, -190643, 82131, -191356, 81854, -192246, 81854
        .AddBezierSegment -192246, 81854, -193148, 81854, -193866, 82129, -194398, 82681
        .AddBezierSegment -194398, 82681, -194930, 83231, -195196, 83990, -195196, 84959
        .CloseSubPath

        .StartSubPath
        .AddLineSegment -193935, 84999
        .AddBezierSegment -193935, 84999, -193935, 84312, -193775, 83791, -193458, 83436
        .AddBezierSegment -193458, 83436, -193141, 83082, -192737, 82904, -192249, 82904
        .AddBezierSegment -192249, 82904, -191759, 82904, -191358, 83080, -191044, 83432
        .AddBezierSegment -191044, 83432, -190730, 83784, -190574, 84312, -190574, 85016
        .AddBezierSegment -190574, 85016, -190574, 85712, -190726, 86231, -191032, 86573
        .AddBezierSegment -191032, 86573, -191337, 86916, -191742, 87087, -192249, 87087
        .AddBezierSegment -192249, 87087, -192754, 87087, -193162, 86913, -193471, 86567
        .AddBezierSegment -193471, 86567, -193780, 86220, -193935, 85698, -193935, 84999
        .CloseSubPath

        .StartSubPath
        .AddLineSegment -186994, 81964
        .AddLineSegment -186994, 87002
        .AddLineSegment -188789, 87002
        .AddLineSegment -188789, 88026
        .AddLineSegment -183980, 88026
        .AddLineSegment -183980, 87002
        .AddLineSegment -185766, 87002
        .AddLineSegment -185766, 81964
        .AddLineSegment -186994, 81964
        .CloseSubPath

        .StartSubPath
        .AddLineSegment -179307, 84199
        .AddLineSegment -178122, 83835
        .AddBezierSegment -178122, 83835, -178304, 83168, -178607, 82672, -179031, 82348
        .AddBezierSegment -179031, 82348, -179454, 82025, -179990, 81862, -180641, 81862
        .AddBezierSegment -180641, 81862, -181447, 81862, -182108, 82138, -182628, 82688
        .AddBezierSegment -182628, 82688, -183146, 83238, -183405, 83990, -183405, 84946
        .AddBezierSegment -183405, 84946, -183405, 85955, -183144, 86739, -182624, 87298
        .AddBezierSegment -182624, 87298, -182103, 87857, -181417, 88136, -180567, 88136
        .AddBezierSegment -180567, 88136, -179825, 88136, -179223, 87916, -178760, 87477
        .AddBezierSegment -178760, 87477, -178483, 87218, -178277, 86845, -178139, 86358
        .AddLineSegment -179350, 86070
        .AddBezierSegment -179350, 86070, -179422, 86384, -179571, 86632, -179798, 86814
        .AddBezierSegment -179798, 86814, -180027, 86996, -180304, 87087, -180628, 87087
        .AddBezierSegment -180628, 87087, -181078, 87087, -181444, 86926, -181723, 86602
        .AddBezierSegment -181723, 86602, -182004, 86281, -182144, 85759, -182144, 85036
        .AddBezierSegment -182144, 85036, -182144, 84271, -182005, 83725, -181730, 83401
        .AddBezierSegment -181730, 83401, -181454, 83075, -181095, 82912, -180654, 82912
        .AddBezierSegment -180654, 82912, -180328, 82912, -180047, 83015, -179813, 83223
        .AddBezierSegment -179813, 83223, -179578, 83430, -179409, 83755, -179307, 84199
        .CloseSubPath

        .StartSubPath
        .AddLineSegment -171610, 81964
        .AddLineSegment -172932, 81964
        .AddLineSegment -173455, 83344
        .AddLineSegment -175887, 83344
        .AddLineSegment -176391, 81964
        .AddLineSegment -177689, 81964
        .AddLineSegment -175335, 88026
        .AddLineSegment -174038, 88026
        .AddLineSegment -171610, 81964
        .CloseSubPath

        .StartSubPath
        .AddLineSegment -173853, 84369
        .AddLineSegment -174696, 86612
        .AddLineSegment -175514, 84369
        .AddLineSegment -173853, 84369
        .CloseSubPath

        .StartSubPath
        .AddLineSegment -170923, 81964
        .AddLineSegment -170923, 87975
        .AddLineSegment -169695, 87975
        .AddLineSegment -169695, 82989
        .AddLineSegment -166655, 82989
        .AddLineSegment -166655, 81964
        .AddLineSegment -170923, 81964
        .CloseSubPath

        .StartSubPath
        .AddLineSegment -161907, 84199
        .AddLineSegment -160721, 83835
        .AddBezierSegment -160721, 83835, -160904, 83168, -161207, 82672, -161630, 82348
        .AddBezierSegment -161630, 82348, -162054, 82025, -162590, 81862, -163240, 81862
        .AddBezierSegment -163240, 81862, -164046, 81862, -164708, 82138, -165227, 82688
        .AddBezierSegment -165227, 82688, -165745, 83238, -166005, 83990, -166005, 84946
        .AddBezierSegment -166005, 84946, -166005, 85955, -165744, 86739, -165223, 87298
        .AddBezierSegment -165223, 87298, -164702, 87857, -164017, 88136, -163167, 88136
        .AddBezierSegment -163167, 88136, -162425, 88136, -161822, 87916, -161359, 87477
        .AddBezierSegment -161359, 87477, -161083, 87218, -160877, 86845, -160738, 86358
        .AddLineSegment -161949, 86070
        .AddBezierSegment -161949, 86070, -162021, 86384, -162171, 86632, -162398, 86814
        .AddBezierSegment -162398, 86814, -162627, 86996, -162903, 87087, -163228, 87087
        .AddBezierSegment -163228, 87087, -163678, 87087, -164043, 86926, -164323, 86602
        .AddBezierSegment -164323, 86602, -164603, 86281, -164743, 85759, -164743, 85036
        .AddBezierSegment -164743, 85036, -164743, 84271, -164605, 83725, -164330, 83401
        .AddBezierSegment -164330, 83401, -164053, 83075, -163695, 82912, -163253, 82912
        .AddBezierSegment -163253, 82912, -162927, 82912, -162646, 83015, -162412, 83223
        .AddBezierSegment -162412, 83223, -162178, 83430, -162008, 83755, -161907, 84199
        .CloseSubPath

        Set .brush = New waBrush
        .brush.SetColor RGB(43, 42, 41), 100
    End With
    MyCurves.Add path

    Set path = New waPath
    With path
        .StartSubPath
        .AddLineSegment -613963, 81964
        .AddLineSegment -613963, 88026
        .AddLineSegment -612000, 88026
        .AddBezierSegment -612000, 88026, -611257, 88026, -610773, 87997, -610547, 87936
        .AddBezierSegment -610547, 87936, -610200, 87844, -609909, 87647, -609675, 87342
        .AddBezierSegment -609675, 87342, -609441, 87037, -609323, 86643, -609323, 86161
        .AddBezierSegment -609323, 86161, -609323, 85788, -609391, 85476, -609525, 85222
        .AddBezierSegment -609525, 85222, -609661, 84968, -609832, 84769, -610039, 84624
        .AddBezierSegment -610039, 84624, -610246, 84480, -610458, 84384, -610673, 84338
        .AddBezierSegment -610673, 84338, -610965, 84280, -611386, 84250, -611940, 84250
        .AddLineSegment -612736, 84250
        .AddLineSegment -612736, 81964
        .AddLineSegment -613963, 81964
        .CloseSubPath

        .StartSubPath
        .AddLineSegment -612736, 87002
        .AddLineSegment -612736, 85275
        .AddLineSegment -612068, 85275
        .AddBezierSegment -612068, 85275, -611587, 85275, -611267, 85307, -611104, 85371
        .AddBezierSegment -611104, 85371, -610942, 85434, -610815, 85533, -610723, 85670
        .AddBezierSegment -610723, 85670, -610632, 85805, -610585, 85962, -610585, 86142
        .AddBezierSegment -610585, 86142, -610585, 86364, -610650, 86546, -610778, 86690
        .AddBezierSegment -610778, 86690, -610908, 86834, -611070, 86924, -611269, 86961
        .AddBezierSegment -611269, 86961, -611415, 86988, -611707, 87002, -612147, 87002
        .AddLineSegment -612736, 87002
        .CloseSubPath

        .StartSubPath
        .AddLineSegment -608282, 81964
        .AddLineSegment -608282, 87975
        .AddLineSegment -607054, 87975
        .AddLineSegment -607054, 82989
        .AddLineSegment -604015, 82989
        .AddLineSegment -604015, 81964
        .AddLineSegment -608282, 81964
        .CloseSubPath

        .StartSubPath
        .AddLineSegment -603398, 84959
        .AddBezierSegment -603398, 84959, -603398, 85577, -603305, 86096, -603120, 86516
        .AddBezierSegment -603120, 86516, -602983, 86825, -602794, 87103, -602556, 87349
        .AddBezierSegment -602556, 87349, -602317, 87594, -602056, 87776, -601773, 87896
        .AddBezierSegment -601773, 87896, -601394, 88056, -600958, 88136, -600465, 88136
        .AddBezierSegment -600465, 88136, -599570, 88136, -598856, 87858, -598318, 87304
        .AddBezierSegment -598318, 87304, -597782, 86748, -597514, 85976, -597514, 84987
        .AddBezierSegment -597514, 84987, -597514, 84006, -597781, 83238, -598313, 82685
        .AddBezierSegment -598313, 82685, -598845, 82131, -599557, 81854, -600448, 81854
        .AddBezierSegment -600448, 81854, -601349, 81854, -602067, 82129, -602600, 82681
        .AddBezierSegment -602600, 82681, -603131, 83231, -603398, 83990, -603398, 84959
        .CloseSubPath

        .StartSubPath
        .AddLineSegment -602137, 84999
        .AddBezierSegment -602137, 84999, -602137, 84312, -601977, 83791, -601660, 83436
        .AddBezierSegment -601660, 83436, -601342, 83082, -600939, 82904, -600450, 82904
        .AddBezierSegment -600450, 82904, -599961, 82904, -599560, 83080, -599245, 83432
        .AddBezierSegment -599245, 83432, -598932, 83784, -598775, 84312, -598775, 85016
        .AddBezierSegment -598775, 85016, -598775, 85712, -598928, 86231, -599234, 86573
        .AddBezierSegment -599234, 86573, -599539, 86916, -599944, 87087, -600450, 87087
        .AddBezierSegment -600450, 87087, -600956, 87087, -601363, 86913, -601672, 86567
        .AddBezierSegment -601672, 86567, -601981, 86220, -602137, 85698, -602137, 84999
        .CloseSubPath

        .StartSubPath
        .AddLineSegment -595195, 81964
        .AddLineSegment -595195, 87002
        .AddLineSegment -596990, 87002
        .AddLineSegment -596990, 88026
        .AddLineSegment -592181, 88026
        .AddLineSegment -592181, 87002
        .AddLineSegment -593968, 87002
        .AddLineSegment -593968, 81964
        .AddLineSegment -595195, 81964
        .CloseSubPath

        .StartSubPath
        .AddLineSegment -587509, 84199
        .AddLineSegment -586324, 83835
        .AddBezierSegment -586324, 83835, -586506, 83168, -586809, 82672, -587232, 82348
        .AddBezierSegment -587232, 82348, -587656, 82025, -588192, 81862, -588843, 81862
        .AddBezierSegment -588843, 81862, -589648, 81862, -590310, 82138, -590829, 82688
        .AddBezierSegment -590829, 82688, -591347, 83238, -591607, 83990, -591607, 84946
        .AddBezierSegment -591607, 84946, -591607, 85955, -591346, 86739, -590825, 87298
        .AddBezierSegment -590825, 87298, -590304, 87857, -589619, 88136, -588769, 88136
        .AddBezierSegment -588769, 88136, -588027, 88136, -587424, 87916, -586962, 87477
        .AddBezierSegment -586962, 87477, -586685, 87218, -586479, 86845, -586341, 86358
        .AddLineSegment -587551, 86070
        .AddBezierSegment -587551, 86070, -587623, 86384, -587773, 86632, -588000, 86814
        .AddBezierSegment -588000, 86814, -588229, 86996, -588505, 87087, -588830, 87087
        .AddBezierSegment -588830, 87087, -589280, 87087, -589645, 86926, -589925, 86602
        .AddBezierSegment -589925, 86602, -590206, 86281, -590345, 85759, -590345, 85036
        .AddBezierSegment -590345, 85036, -590345, 84271, -590207, 83725, -589932, 83401
        .AddBezierSegment -589932, 83401, -589655, 83075, -589297, 82912, -588855, 82912
        .AddBezierSegment -588855, 82912, -588529, 82912, -588249, 83015, -588014, 83223
        .AddBezierSegment -588014, 83223, -587780, 83430, -587611, 83755, -587509, 84199
        .CloseSubPath

        .StartSubPath
        .AddLineSegment -579811, 81964
        .AddLineSegment -581134, 81964
        .AddLineSegment -581657, 83344
        .AddLineSegment -584088, 83344
        .AddLineSegment -584592, 81964
        .AddLineSegment -585890, 81964
        .AddLineSegment -583537, 88026
        .AddLineSegment -582240, 88026
        .AddLineSegment -579811, 81964
        .CloseSubPath

        .StartSubPath
        .AddLineSegment -582055, 84369
        .AddLineSegment -582898, 86612
        .AddLineSegment -583716, 84369
        .AddLineSegment -582055, 84369
        .CloseSubPath

        .StartSubPath
        .AddLineSegment -579124, 81964
        .AddLineSegment -579124, 87975
        .AddLineSegment -577896, 87975
        .AddLineSegment -577896, 82989
        .AddLineSegment -574857, 82989
        .AddLineSegment -574857, 81964
        .AddLineSegment -579124, 81964
        .CloseSubPath

        .StartSubPath
        .AddLineSegment -570109, 84199
        .AddLineSegment -568923, 83835
        .AddBezierSegment -568923, 83835, -569105, 83168, -569409, 82672, -569832, 82348
        .AddBezierSegment -569832, 82348, -570255, 82025, -570792, 81862, -571442, 81862
        .AddBezierSegment -571442, 81862, -572248, 81862, -572910, 82138, -573429, 82688
        .AddBezierSegment -573429, 82688, -573947, 83238, -574207, 83990, -574207, 84946
        .AddBezierSegment -574207, 84946, -574207, 85955, -573945, 86739, -573425, 87298
        .AddBezierSegment -573425, 87298, -572904, 87857, -572218, 88136, -571369, 88136
        .AddBezierSegment -571369, 88136, -570626, 88136, -570024, 87916, -569561, 87477
        .AddBezierSegment -569561, 87477, -569284, 87218, -569078, 86845, -568940, 86358
        .AddLineSegment -570151, 86070
        .AddBezierSegment -570151, 86070, -570223, 86384, -570372, 86632, -570600, 86814
        .AddBezierSegment -570600, 86814, -570828, 86996, -571105, 87087, -571429, 87087
        .AddBezierSegment -571429, 87087, -571880, 87087, -572245, 86926, -572524, 86602
        .AddBezierSegment -572524, 86602, -572805, 86281, -572945, 85759, -572945, 85036
        .AddBezierSegment -572945, 85036, -572945, 84271, -572807, 83725, -572532, 83401
        .AddBezierSegment -572532, 83401, -572255, 83075, -571896, 82912, -571455, 82912
        .AddBezierSegment -571455, 82912, -571129, 82912, -570848, 83015, -570614, 83223
        .AddBezierSegment -570614, 83223, -570380, 83430, -570210, 83755, -570109, 84199
        .CloseSubPath

        Set .brush = New waBrush
        .brush.SetColor RGB(43, 42, 41), 100
    End With
    MyCurves.Add path
End Sub

' -----------------------------------------------------------------------------
' InitAxisCurves (Sub), source lines 9678-10006
' -----------------------------------------------------------------------------
Private Sub InitAxisCurves()
    Dim MyCurves As New Collection
    Set AxisCurves = MyCurves
    Dim path As waPath

    Set path = New waPath
    With path
        .StartSubPath
        .AddLineSegment 3321, -1015
        .AddBezierSegment 3321, -1015, 3648, -1115, 3935, -1307, 4163, -1596
        .AddBezierSegment 4163, -1596, 4618, -2171, 4710, -3024, 4710, -3732
        .AddBezierSegment 4710, -3732, 4710, -4443, 4617, -5295, 4160, -5872
        .AddBezierSegment 4160, -5872, 3805, -6321, 3305, -6530, 2737, -6530
        .AddBezierSegment 2737, -6530, 2147, -6530, 1643, -6292, 1276, -5830
        .AddBezierSegment 1276, -5830, 1134, -5652, 1031, -5442, 956, -5215
        .AddLineSegment 956, -5670
        .AddBezierSegment 956, -5670, 956, -5911, 940, -6173, 871, -6404
        .AddBezierSegment 871, -6404, 800, -6642, 667, -6849, 488, -7021
        .AddBezierSegment 488, -7021, 306, -7195, 87, -7319, -149, -7404
        .AddBezierSegment -149, -7404, -412, -7499, -702, -7379, -821, -7125
        .AddLineSegment -985, -6774
        .AddBezierSegment -985, -6774, -1053, -6629, -1056, -6478, -993, -6332
        .AddBezierSegment -993, -6332, -968, -6274, -935, -6223, -896, -6179
        .AddBezierSegment -896, -6179, -942, -6100, -969, -6008, -969, -5910
        .AddLineSegment -969, -5198
        .AddBezierSegment -969, -5198, -1051, -5444, -1164, -5674, -1321, -5872
        .AddBezierSegment -1321, -5872, -1676, -6321, -2175, -6530, -2744, -6530
        .AddBezierSegment -2744, -6530, -3335, -6530, -3838, -6293, -4206, -5830
        .AddBezierSegment -4206, -5830, -4645, -5278, -4717, -4402, -4717, -3724
        .AddBezierSegment -4717, -3724, -4717, -3013, -4623, -2166, -4168, -1590
        .AddBezierSegment -4168, -1590, -3812, -1140, -3313, -932, -2744, -932
        .AddBezierSegment -2744, -932, -2173, -932, -1675, -1145, -1318, -1596
        .AddBezierSegment -1318, -1596, -1121, -1845, -992, -2146, -909, -2465
        .AddBezierSegment -909, -2465, -820, -2290, -637, -2171, -428, -2171
        .AddLineSegment 416, -2171
        .AddBezierSegment 416, -2171, 627, -2171, 811, -2293, 900, -2471
        .AddBezierSegment 900, -2471, 983, -2148, 1113, -1843, 1314, -1590
        .AddBezierSegment 1314, -1590, 1581, -1252, 1929, -1050, 2327, -971
        .AddBezierSegment 2327, -971, 2253, -938, 2187, -888, 2131, -821
        .AddBezierSegment 2131, -821, 2025, -693, 1985, -540, 2016, -377
        .AddBezierSegment 2016, -377, 2039, -253, 2050, -126, 2050, 0
        .AddBezierSegment 2050, 0, 2050, 126, 2039, 253, 2016, 377
        .AddBezierSegment 2016, 377, 1985, 540, 2025, 693, 2131, 821
        .AddBezierSegment 2131, 821, 2237, 948, 2380, 1015, 2546, 1015
        .AddLineSegment 38241, 1015
        .AddLineSegment 38241, 1613
        .AddBezierSegment 38241, 1613, 38241, 1794, 38321, 1949, 38469, 2054
        .AddBezierSegment 38469, 2054, 38616, 2158, 38790, 2182, 38960, 2123
        .AddLineSegment 43550, 509
        .AddBezierSegment 43550, 509, 43770, 432, 43910, 233, 43910, 0
        .AddBezierSegment 43910, 0, 43910, -233, 43770, -432, 43550, -509
        .AddLineSegment 40000, -1757
        .AddBezierSegment 40000, -1757, 40043, -1788, 40081, -1827, 40114, -1872
        .AddLineSegment 41085, -3179
        .AddLineSegment 42080, -1868
        .AddBezierSegment 42080, -1868, 42186, -1729, 42336, -1654, 42510, -1654
        .AddLineSegment 43353, -1654
        .AddBezierSegment 43353, -1654, 43564, -1654, 43745, -1767, 43838, -1957
        .AddBezierSegment 43838, -1957, 43931, -2146, 43909, -2358, 43780, -2525
        .AddLineSegment 42190, -4573
        .AddLineSegment 43784, -6686
        .AddBezierSegment 43784, -6686, 43911, -6853, 43930, -7064, 43837, -7252
        .AddBezierSegment 43837, -7252, 43743, -7439, 43563, -7551, 43353, -7551
        .AddLineSegment 42461, -7551
        .AddBezierSegment 42461, -7551, 42285, -7551, 42133, -7475, 42028, -7334
        .AddLineSegment 41026, -5993
        .AddLineSegment 40005, -7337
        .AddBezierSegment 40005, -7337, 39899, -7476, 39749, -7551, 39575, -7551
        .AddLineSegment 38732, -7551
        .AddBezierSegment 38732, -7551, 38521, -7551, 38340, -7438, 38247, -7248
        .AddBezierSegment 38247, -7248, 38154, -7058, 38176, -6846, 38305, -6679
        .AddLineSegment 39924, -4596
        .AddLineSegment 38357, -2519
        .AddBezierSegment 38357, -2519, 38231, -2352, 38212, -2141, 38305, -1953
        .AddBezierSegment 38305, -1953, 38312, -1939, 38320, -1925, 38328, -1911
        .AddBezierSegment 38328, -1911, 38271, -1825, 38241, -1724, 38241, -1613
        .AddLineSegment 38241, -1015
        .AddLineSegment 3321, -1015
        .CloseSubPath

        .StartSubPath
        .AddLineSegment 837, -4728
        .AddBezierSegment 837, -4728, 794, -4472, 774, -4210, 767, -3965
        .AddBezierSegment 767, -3965, 672, -4046, 550, -4095, 416, -4095
        .AddLineSegment -428, -4095
        .AddBezierSegment -428, -4095, -560, -4095, -681, -4048, -775, -3969
        .AddBezierSegment -775, -3969, -783, -4217, -804, -4476, -851, -4729
        .AddBezierSegment -851, -4729, -751, -4605, -599, -4525, -428, -4525
        .AddLineSegment 416, -4525
        .AddBezierSegment 416, -4525, 586, -4525, 738, -4605, 837, -4728
        .CloseSubPath

        .StartSubPath
        .AddLineSegment -672, 43066
        .AddBezierSegment -672, 43066, -675, 43061, -678, 43056, -682, 43051
        .AddLineSegment -2560, 40510
        .AddLineSegment -2560, 38556
        .AddBezierSegment -2560, 38556, -2560, 38258, -2803, 38016, -3100, 38016
        .AddLineSegment -3882, 38016
        .AddBezierSegment -3882, 38016, -4180, 38016, -4423, 38258, -4423, 38556
        .AddLineSegment -4423, 40445
        .AddLineSegment -6305, 43057
        .AddBezierSegment -6305, 43057, -6426, 43225, -6442, 43433, -6347, 43618
        .AddBezierSegment -6347, 43618, -6253, 43803, -6074, 43912, -5867, 43912
        .AddLineSegment -5003, 43912
        .AddBezierSegment -5003, 43912, -4824, 43912, -4669, 43833, -4564, 43686
        .AddLineSegment -3484, 42173
        .AddLineSegment -2384, 43689
        .AddBezierSegment -2384, 43689, -2279, 43834, -2126, 43912, -1947, 43912
        .AddLineSegment -1116, 43912
        .AddBezierSegment -1116, 43912, -907, 43912, -728, 43802, -634, 43615
        .AddBezierSegment -634, 43615, -592, 43532, -572, 43443, -575, 43355
        .AddLineSegment -512, 43542
        .AddBezierSegment -512, 43542, -437, 43766, -236, 43910, 0, 43910
        .AddBezierSegment 0, 43910, 236, 43910, 437, 43766, 512, 43542
        .AddLineSegment 2125, 38729
        .AddBezierSegment 2125, 38729, 2182, 38559, 2156, 38387, 2051, 38242
        .AddBezierSegment 2051, 38242, 1947, 38096, 1792, 38017, 1613, 38017
        .AddLineSegment 1015, 38017
        .AddLineSegment 1015, 2546
        .AddBezierSegment 1015, 2546, 1015, 2380, 948, 2237, 821, 2131
        .AddBezierSegment 821, 2131, 693, 2025, 540, 1985, 377, 2016
        .AddBezierSegment 377, 2016, 253, 2039, 126, 2050, 0, 2050
        .AddBezierSegment 0, 2050, -126, 2050, -253, 2039, -377, 2016
        .AddBezierSegment -377, 2016, -540, 1985, -693, 2025, -821, 2131
        .AddBezierSegment -821, 2131, -948, 2237, -1015, 2380, -1015, 2546
        .AddLineSegment -1015, 38017
        .AddLineSegment -1613, 38017
        .AddBezierSegment -1613, 38017, -1792, 38017, -1947, 38096, -2051, 38242
        .AddBezierSegment -2051, 38242, -2156, 38387, -2182, 38559, -2125, 38729
        .AddLineSegment -672, 43066
        .CloseSubPath

        .StartSubPath
        .AddLineSegment 0, -1350
        .AddBezierSegment 0, -1350, -745, -1350, -1350, -745, -1350, 0
        .AddBezierSegment -1350, 0, -1350, 745, -745, 1350, 0, 1350
        .AddBezierSegment 0, 1350, 745, 1350, 1350, 745, 1350, 0
        .AddBezierSegment 1350, 0, 1350, -745, 745, -1350, 0, -1350
        .CloseSubPath

        Set .brush = New waBrush
        .brush.SetColor RGB(254, 254, 254), 75
'        .brush.SetColor RGB(0, 0, 0), 20
    End With
    MyCurves.Add path

    Set path = New waPath
    With path
        .StartSubPath
        .AddLineSegment 43353, -7011
        .AddLineSegment 42461, -7011
        .AddLineSegment 41029, -5096
        .AddLineSegment 39575, -7011
        .AddLineSegment 38732, -7011
        .AddLineSegment 40604, -4600
        .AddLineSegment 38788, -2194
        .AddLineSegment 39681, -2194
        .AddLineSegment 41080, -4079
        .AddLineSegment 42510, -2194
        .AddLineSegment 43353, -2194
        .AddLineSegment 41510, -4568
        .AddLineSegment 43353, -7011
        .CloseSubPath

        Set .brush = New waBrush
        .brush.SetColor RGB(227, 30, 36), 80
    End With
    MyCurves.Add path

    Set path = New waPath
    With path
        .StartSubPath
        .AddLineSegment 475, 2546
        .AddLineSegment 475, 38557
        .AddLineSegment 1613, 38557
        .AddLineSegment 0, 43370
        .AddLineSegment -1613, 38557
        .AddLineSegment -475, 38557
        .AddLineSegment -475, 2546
        .AddBezierSegment -475, 2546, -321, 2575, -162, 2590, 0, 2590
        .AddBezierSegment 0, 2590, 162, 2590, 321, 2575, 475, 2546
        .CloseSubPath

        Set .brush = New waBrush
        .brush.SetColor RGB(0, 122, 70), 80
    End With
    MyCurves.Add path

    Set path = New waPath
    With path
        .StartSubPath
        .AddLineSegment 2546, -475
        .AddLineSegment 38781, -475
        .AddLineSegment 38781, -1613
        .AddLineSegment 43370, 0
        .AddLineSegment 38781, 1613
        .AddLineSegment 38781, 475
        .AddLineSegment 2546, 475
        .AddBezierSegment 2546, 475, 2575, 321, 2590, 162, 2590, 0
        .AddBezierSegment 2590, 0, 2590, -162, 2575, -321, 2546, -475
        .CloseSubPath

        Set .brush = New waBrush
        .brush.SetColor RGB(227, 30, 36), 80
    End With
    MyCurves.Add path

    Set path = New waPath
    With path
        .StartSubPath
        .AddLineSegment 0, -810
        .AddBezierSegment 0, -810, 447, -810, 810, -447, 810, 0
        .AddBezierSegment 810, 0, 810, 447, 447, 810, 0, 810
        .AddBezierSegment 0, 810, -447, 810, -810, 447, -810, 0
        .AddBezierSegment -810, 0, -810, -447, -447, -810, 0, -810
        .CloseSubPath

        Set .brush = New waBrush
        .brush.SetColor RGB(43, 42, 41), 80
    End With
    MyCurves.Add path

    Set path = New waPath
    With path
        .StartSubPath
        .AddLineSegment -1116, 43372
        .AddLineSegment -3100, 40688
        .AddLineSegment -3100, 38556
        .AddLineSegment -3882, 38556
        .AddLineSegment -3882, 40620
        .AddLineSegment -5867, 43372
        .AddLineSegment -5003, 43372
        .AddLineSegment -3487, 41249
        .AddLineSegment -1947, 43372
        .AddLineSegment -1116, 43372
        .CloseSubPath

        Set .brush = New waBrush
        .brush.SetColor RGB(0, 122, 70), 80
    End With
    MyCurves.Add path

    Set path = New waPath
    With path
        .StartSubPath
        .AddLineSegment -2744, -1472
        .AddBezierSegment -2744, -1472, -2317, -1472, -1983, -1625, -1742, -1931
        .AddBezierSegment -1742, -1931, -1455, -2292, -1311, -2893, -1311, -3732
        .AddBezierSegment -1311, -3732, -1311, -4569, -1456, -5171, -1745, -5537
        .AddBezierSegment -1745, -5537, -1984, -5839, -2317, -5990, -2744, -5990
        .AddBezierSegment -2744, -5990, -3174, -5990, -3520, -5825, -3783, -5494
        .AddBezierSegment -3783, -5494, -4046, -5163, -4177, -4572, -4177, -3724
        .AddBezierSegment -4177, -3724, -4177, -2890, -4033, -2290, -3744, -1925
        .AddBezierSegment -3744, -1925, -3505, -1623, -3172, -1472, -2744, -1472
        .CloseSubPath

        .StartSubPath
        .AddLineSegment -2744, -2174
        .AddBezierSegment -2744, -2174, -2845, -2174, -2935, -2207, -3013, -2273
        .AddBezierSegment -3013, -2273, -3092, -2338, -3154, -2455, -3197, -2624
        .AddBezierSegment -3197, -2624, -3255, -2843, -3284, -3213, -3284, -3732
        .AddBezierSegment -3284, -3732, -3284, -4252, -3258, -4608, -3207, -4802
        .AddBezierSegment -3207, -4802, -3155, -4997, -3091, -5126, -3012, -5191
        .AddBezierSegment -3012, -5191, -2934, -5255, -2845, -5287, -2744, -5287
        .AddBezierSegment -2744, -5287, -2643, -5287, -2553, -5254, -2474, -5189
        .AddBezierSegment -2474, -5189, -2396, -5124, -2334, -5007, -2291, -4837
        .AddBezierSegment -2291, -4837, -2234, -4620, -2205, -4252, -2205, -3732
        .AddBezierSegment -2205, -3732, -2205, -3213, -2231, -2857, -2282, -2662
        .AddBezierSegment -2282, -2662, -2333, -2468, -2398, -2338, -2476, -2273
        .AddBezierSegment -2476, -2273, -2554, -2207, -2643, -2174, -2744, -2174
        .CloseSubPath

        .StartSubPath
        .AddLineSegment -428, -3555
        .AddLineSegment -428, -2711
        .AddLineSegment 416, -2711
        .AddLineSegment 416, -3555
        .AddLineSegment -428, -3555
        .CloseSubPath

        .StartSubPath
        .AddLineSegment -428, -5065
        .AddLineSegment 416, -5065
        .AddLineSegment 416, -5670
        .AddBezierSegment 416, -5670, 416, -5915, 396, -6109, 353, -6250
        .AddBezierSegment 353, -6250, 311, -6392, 231, -6519, 114, -6631
        .AddBezierSegment 114, -6631, -3, -6744, -152, -6831, -332, -6896
        .AddLineSegment -496, -6545
        .AddBezierSegment -496, -6545, -327, -6488, -207, -6410, -135, -6311
        .AddBezierSegment -135, -6311, -64, -6210, -26, -6076, -22, -5910
        .AddLineSegment -428, -5910
        .AddLineSegment -428, -5065
        .CloseSubPath

        .StartSubPath
        .AddLineSegment 2737, -1472
        .AddBezierSegment 2737, -1472, 3164, -1472, 3498, -1625, 3740, -1931
        .AddBezierSegment 3740, -1931, 4026, -2292, 4170, -2893, 4170, -3732
        .AddBezierSegment 4170, -3732, 4170, -4569, 4025, -5171, 3736, -5537
        .AddBezierSegment 3736, -5537, 3497, -5839, 3164, -5990, 2737, -5990
        .AddBezierSegment 2737, -5990, 2308, -5990, 1961, -5825, 1698, -5494
        .AddBezierSegment 1698, -5494, 1435, -5163, 1304, -4572, 1304, -3724
        .AddBezierSegment 1304, -3724, 1304, -2890, 1449, -2290, 1737, -1925
        .AddBezierSegment 1737, -1925, 1977, -1623, 2309, -1472, 2737, -1472
        .CloseSubPath

        .StartSubPath
        .AddLineSegment 2737, -2174
        .AddBezierSegment 2737, -2174, 2636, -2174, 2546, -2207, 2468, -2273
        .AddBezierSegment 2468, -2273, 2389, -2338, 2327, -2455, 2284, -2624
        .AddBezierSegment 2284, -2624, 2226, -2843, 2197, -3213, 2197, -3732
        .AddBezierSegment 2197, -3732, 2197, -4252, 2223, -4608, 2275, -4802
        .AddBezierSegment 2275, -4802, 2326, -4997, 2391, -5126, 2469, -5191
        .AddBezierSegment 2469, -5191, 2547, -5255, 2636, -5287, 2737, -5287
        .AddBezierSegment 2737, -5287, 2838, -5287, 2928, -5254, 3007, -5189
        .AddBezierSegment 3007, -5189, 3085, -5124, 3147, -5007, 3190, -4837
        .AddBezierSegment 3190, -4837, 3248, -4620, 3276, -4252, 3276, -3732
        .AddBezierSegment 3276, -3732, 3276, -3213, 3251, -2857, 3199, -2662
        .AddBezierSegment 3199, -2662, 3148, -2468, 3083, -2338, 3005, -2273
        .AddBezierSegment 3005, -2273, 2927, -2207, 2838, -2174, 2737, -2174
        .CloseSubPath

        Set .brush = New waBrush
        .brush.SetColor RGB(43, 42, 41), 80
    End With
    MyCurves.Add path
    AxisCurves_Initialized = True
End Sub

' -----------------------------------------------------------------------------
' InitLogoCurves (Sub), source lines 10007-10225
' -----------------------------------------------------------------------------
Private Sub InitLogoCurves()
    Dim MyCurves As New Collection
    Set LogoCurves = MyCurves
    Dim path As waPath

    Set path = New waPath
    With path
        .StartSubPath
        .AddLineSegment 275705, -4829
        .AddLineSegment 260693, -7497
        .AddBezierSegment 260693, -7497, 260193, -4513, 259044, -2271, 257246, -751
        .AddBezierSegment 257246, -751, 255449, 750, 253132, 1510, 250278, 1510
        .AddBezierSegment 250278, 1510, 246479, 1510, 243439, 194, 241178, -2438
        .AddBezierSegment 241178, -2438, 238917, -5069, 237787, -9462, 237787, -15633
        .AddBezierSegment 237787, -15633, 237787, -22491, 238936, -27328, 241234, -30163
        .AddBezierSegment 241234, -30163, 243532, -32999, 246608, -34407, 250463, -34407
        .AddBezierSegment 250463, -34407, 253355, -34407, 255727, -33573, 257561, -31924
        .AddBezierSegment 257561, -31924, 259415, -30274, 260712, -27439, 261472, -23399
        .AddLineSegment 276484, -25956
        .AddBezierSegment 276484, -25956, 274927, -32851, 271943, -38058, 267532, -41580
        .AddBezierSegment 267532, -41580, 263121, -45101, 257209, -46862, 249778, -46862
        .AddBezierSegment 249778, -46862, 241363, -46862, 234636, -44193, 229632, -38874
        .AddBezierSegment 229632, -38874, 224609, -33536, 222107, -26160, 222107, -16727
        .AddBezierSegment 222107, -16727, 222107, -7182, 224628, 231, 229650, 5550
        .AddBezierSegment 229650, 5550, 234673, 10869, 241475, 13519, 250056, 13519
        .AddBezierSegment 250056, 13519, 257061, 13519, 262640, 12000, 266791, 8979
        .AddBezierSegment 266791, 8979, 270924, 5958, 273889, 1362, 275705, -4829
        .CloseSubPath

        .StartSubPath
        .AddLineSegment 194549, -45416
        .AddLineSegment 194549, 34202
        .AddLineSegment 209783, 34202
        .AddLineSegment 209783, -45416
        .AddLineSegment 194549, -45416
        .CloseSubPath

        .StartSubPath
        .AddLineSegment 164506, -17839
        .AddBezierSegment 164506, -17839, 162560, -18487, 159465, -19266, 155221, -20174
        .AddBezierSegment 155221, -20174, 150996, -21082, 148216, -21972, 146918, -22824
        .AddBezierSegment 146918, -22824, 144935, -24251, 143935, -26030, 143935, -28199
        .AddBezierSegment 143935, -28199, 143935, -30349, 144732, -32183, 146325, -33740
        .AddBezierSegment 146325, -33740, 147919, -35297, 149939, -36075, 152404, -36075
        .AddBezierSegment 152404, -36075, 155147, -36075, 157779, -35167, 160281, -33370
        .AddBezierSegment 160281, -33370, 162116, -31980, 163339, -30312, 163913, -28310
        .AddBezierSegment 163913, -28310, 164303, -27013, 164506, -24529, 164506, -20878
        .AddLineSegment 164506, -17839
        .CloseSubPath

        .StartSubPath
        .AddLineSegment 144046, -5496
        .AddLineSegment 130257, -3049
        .AddBezierSegment 130257, -3049, 131814, 2585, 134501, 6755, 138301, 9460
        .AddBezierSegment 138301, 9460, 142100, 12166, 147752, 13519, 155240, 13519
        .AddBezierSegment 155240, 13519, 162042, 13519, 167120, 12722, 170456, 11110
        .AddBezierSegment 170456, 11110, 173773, 9498, 176127, 7459, 177480, 4976
        .AddBezierSegment 177480, 4976, 178833, 2492, 179518, -2049, 179518, -8683
        .AddLineSegment 179296, -26475
        .AddBezierSegment 179296, -26475, 179296, -31535, 179537, -35279, 180019, -37688
        .AddBezierSegment 180019, -37688, 180500, -40097, 181409, -42673, 182743, -45416
        .AddLineSegment 167620, -45416
        .AddBezierSegment 167620, -45416, 167231, -44415, 166730, -42933, 166156, -40968
        .AddBezierSegment 166156, -40968, 165915, -40060, 165730, -39467, 165618, -39189
        .AddBezierSegment 165618, -39189, 163024, -41746, 160244, -43655, 157279, -44934
        .AddBezierSegment 157279, -44934, 154313, -46213, 151163, -46862, 147808, -46862
        .AddBezierSegment 147808, -46862, 141896, -46862, 137226, -45249, 133816, -42025
        .AddBezierSegment 133816, -42025, 130405, -38818, 128700, -34741, 128700, -29811
        .AddBezierSegment 128700, -29811, 128700, -26568, 129479, -23658, 131036, -21101
        .AddBezierSegment 131036, -21101, 132592, -18562, 134761, -16597, 137578, -15244
        .AddBezierSegment 137578, -15244, 140376, -13891, 144417, -12705, 149698, -11686
        .AddBezierSegment 149698, -11686, 156815, -10351, 161764, -9091, 164506, -7942
        .AddLineSegment 164506, -6404
        .AddBezierSegment 164506, -6404, 164506, -3438, 163784, -1326, 162338, -47
        .AddBezierSegment 162338, -47, 160874, 1213, 158131, 1843, 154109, 1843
        .AddBezierSegment 154109, 1843, 151385, 1843, 149272, 1306, 147752, 212
        .AddBezierSegment 147752, 212, 146214, -862, 144991, -2771, 144046, -5496
        .CloseSubPath

        .StartSubPath
        .AddLineSegment 103440, -16060
        .AddLineSegment 119008, -20841
        .AddBezierSegment 119008, -20841, 116617, -29607, 112632, -36112, 107072, -40375
        .AddBezierSegment 107072, -40375, 101512, -44619, 94470, -46751, 85926, -46751
        .AddBezierSegment 85926, -46751, 75344, -46751, 66651, -43136, 59831, -35909
        .AddBezierSegment 59831, -35909, 53029, -28681, 49619, -18802, 49619, -6255
        .AddBezierSegment 49619, -6255, 49619, 6996, 53048, 17300, 59887, 24639
        .AddBezierSegment 59887, 24639, 66725, 31978, 75733, 35648, 86890, 35648
        .AddBezierSegment 86890, 35648, 96638, 35648, 104552, 32757, 110630, 26993
        .AddBezierSegment 110630, 26993, 114263, 23583, 116969, 18690, 118785, 12296
        .AddLineSegment 102884, 8515
        .AddBezierSegment 102884, 8515, 101939, 12630, 99974, 15892, 96990, 18282
        .AddBezierSegment 96990, 18282, 93988, 20673, 90355, 21859, 86093, 21859
        .AddBezierSegment 86093, 21859, 80181, 21859, 75380, 19747, 71711, 15502
        .AddBezierSegment 71711, 15502, 68023, 11277, 66188, 4420, 66188, -5069
        .AddBezierSegment 66188, -5069, 66188, -15114, 68004, -22287, 71618, -26549
        .AddBezierSegment 71618, -26549, 75251, -30831, 79958, -32962, 85759, -32962
        .AddBezierSegment 85759, -32962, 90040, -32962, 93728, -31609, 96805, -28884
        .AddBezierSegment 96805, -28884, 99881, -26160, 102105, -21898, 103440, -16060
        .CloseSubPath

        .StartSubPath
        .AddLineSegment 41724, 12296
        .AddLineSegment 41724, 175
        .AddLineSegment 31271, 175
        .AddLineSegment 31271, -23195
        .AddBezierSegment 31271, -23195, 31271, -27939, 31364, -30701, 31568, -31479
        .AddBezierSegment 31568, -31479, 31772, -32258, 32217, -32906, 32940, -33425
        .AddBezierSegment 32940, -33425, 33644, -33926, 34496, -34185, 35516, -34185
        .AddBezierSegment 35516, -34185, 36943, -34185, 38981, -33703, 41669, -32721
        .AddLineSegment 42947, -44582
        .AddBezierSegment 42947, -44582, 39408, -46102, 35404, -46862, 30919, -46862
        .AddBezierSegment 30919, -46862, 28177, -46862, 25712, -46398, 23506, -45472
        .AddBezierSegment 23506, -45472, 21301, -44545, 19688, -43340, 18669, -41876
        .AddBezierSegment 18669, -41876, 17631, -40394, 16927, -38411, 16519, -35890
        .AddBezierSegment 16519, -35890, 16204, -34111, 16037, -30515, 16037, -25085
        .AddLineSegment 16037, 175
        .AddLineSegment 9032, 175
        .AddLineSegment 9032, 12296
        .AddLineSegment 16037, 12296
        .AddLineSegment 16037, 23750
        .AddLineSegment 31271, 32757
        .AddLineSegment 31271, 12296
        .AddLineSegment 41724, 12296
        .CloseSubPath

        .StartSubPath
        .AddLineSegment -40433, -16671
        .AddBezierSegment -40433, -16671, -40433, -22398, -39080, -26790, -36374, -29830
        .AddBezierSegment -36374, -29830, -33650, -32888, -30314, -34407, -26329, -34407
        .AddBezierSegment -26329, -34407, -22363, -34407, -19027, -32888, -16340, -29830
        .AddBezierSegment -16340, -29830, -13653, -26790, -12300, -22361, -12300, -16560
        .AddBezierSegment -12300, -16560, -12300, -10907, -13653, -6552, -16340, -3513
        .AddBezierSegment -16340, -3513, -19027, -455, -22363, 1065, -26329, 1065
        .AddBezierSegment -26329, 1065, -30314, 1065, -33650, -455, -36374, -3513
        .AddBezierSegment -36374, -3513, -39080, -6552, -40433, -10944, -40433, -16671
        .CloseSubPath

        .StartSubPath
        .AddLineSegment -56112, -15856
        .AddBezierSegment -56112, -15856, -56112, -10778, -54871, -5866, -52369, -1122
        .AddBezierSegment -52369, -1122, -49867, 3641, -46327, 7255, -41749, 9757
        .AddBezierSegment -41749, 9757, -37171, 12259, -32075, 13519, -26422, 13519
        .AddBezierSegment -26422, 13519, -17693, 13519, -10558, 10684, -4979, 5013
        .AddBezierSegment -4979, 5013, 599, -677, 3379, -7831, 3379, -16504
        .AddBezierSegment 3379, -16504, 3379, -25252, 562, -32498, -5053, -38244
        .AddBezierSegment -5053, -38244, -10688, -43989, -17767, -46862, -26311, -46862
        .AddBezierSegment -26311, -46862, -31593, -46862, -36634, -45657, -41434, -43266
        .AddBezierSegment -41434, -43266, -46216, -40876, -49867, -37373, -52369, -32739
        .AddBezierSegment -52369, -32739, -54871, -28125, -56112, -22491, -56112, -15856
        .CloseSubPath

        .StartSubPath
        .AddLineSegment -83449, -45416
        .AddLineSegment -83449, 34202
        .AddLineSegment -68215, 34202
        .AddLineSegment -68215, -45416
        .AddLineSegment -83449, -45416
        .CloseSubPath

        .StartSubPath
        .AddLineSegment -141383, 20747
        .AddLineSegment -141383, -1937
        .AddLineSegment -132617, -1937
        .AddBezierSegment -132617, -1937, -126297, -1937, -122090, -1511, -119959, -677
        .AddBezierSegment -119959, -677, -117828, 157, -116160, 1454, -114955, 3252
        .AddBezierSegment -114955, 3252, -113751, 5031, -113139, 7088, -113139, 9460
        .AddBezierSegment -113139, 9460, -113139, 12370, -113991, 14761, -115678, 16651
        .AddBezierSegment -115678, 16651, -117383, 18542, -119514, 19728, -122127, 20210
        .AddBezierSegment -122127, 20210, -124036, 20562, -127873, 20747, -133655, 20747
        .AddLineSegment -141383, 20747
        .CloseSubPath

        .StartSubPath
        .AddLineSegment -157507, -45416
        .AddLineSegment -157507, 34202
        .AddLineSegment -131728, 34202
        .AddBezierSegment -131728, 34202, -121961, 34202, -115604, 33813, -112638, 33016
        .AddBezierSegment -112638, 33016, -108079, 31812, -104261, 29217, -101185, 25214
        .AddBezierSegment -101185, 25214, -98108, 21211, -96570, 16040, -96570, 9701
        .AddBezierSegment -96570, 9701, -96570, 4809, -97460, 713, -99221, -2623
        .AddBezierSegment -99221, -2623, -101000, -5959, -103242, -8572, -105966, -10481
        .AddBezierSegment -105966, -10481, -108691, -12371, -111471, -13632, -114288, -14243
        .AddBezierSegment -114288, -14243, -118124, -15003, -123666, -15392, -130931, -15392
        .AddLineSegment -141383, -15392
        .AddLineSegment -141383, -45416
        .AddLineSegment -157507, -45416
        .CloseSubPath
        
    End With
    MyCurves.Add path

    Set path = New waPath
    With path
        .StartSubPath
        .AddLineSegment -276484, 24505
        .AddLineSegment -249645, 24505
        .AddBezierSegment -249645, 24505, -249551, 30784, -260153, 46837, -240849, 46862
        .AddBezierSegment -240849, 46862, -221433, 46886, -233323, 30106, -233405, 24505
        .AddLineSegment -206566, 24505
        .AddLineSegment -206566, -983
        .AddBezierSegment -206566, -983, -200992, -983, -187643, 10001, -187593, -9329
        .AddBezierSegment -187593, -9329, -187542, -29286, -201587, -17223, -206566, -17223
        .AddLineSegment -206566, -45416
        .AddLineSegment -230925, -45416
        .AddBezierSegment -230925, -45416, -230925, -37445, -219873, -18550, -239043, -18550
        .AddBezierSegment -239043, -18550, -258214, -18550, -247163, -37444, -247163, -45416
        .AddLineSegment -276484, -45416
        .AddLineSegment -274679, -12629
        .AddBezierSegment -274679, -12629, -263630, -12608, -270765, -16013, -262050, -16013
        .AddBezierSegment -262050, -16013, -246275, -16013, -246275, 7949, -262050, 7949
        .AddBezierSegment -262050, 7949, -269164, 7949, -266939, 3676, -274442, 3728
        .AddLineSegment -276484, 24505
        .CloseSubPath
        
    End With
    MyCurves.Add path
    LogoCurvesInitialized = True
End Sub
