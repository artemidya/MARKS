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
