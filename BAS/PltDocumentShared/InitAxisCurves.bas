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
