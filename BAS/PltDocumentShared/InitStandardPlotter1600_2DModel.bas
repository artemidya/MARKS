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
