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
