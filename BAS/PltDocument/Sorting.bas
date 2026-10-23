Private Sub Sorting()
    'Step 1 Sort
    SortCurves
    
    'Step 2
    Dim q As Long
    Dim MyCurves As New Collection
    Dim MySortingCurve As SortingCurve
    
    For q = 1 To Curves.count
        Set MySortingCurve = New SortingCurve
        With MySortingCurve
            Set .Childs = New Collection
            Set .path = Curves(q)
        End With
        MyCurves.Add MySortingCurve
    Next
    
    'Step 3 IsInner
    SetChilds MyCurves
    
    'Step 4 Set Bloks
    Dim FilmBlocks() As FilmBlock
    Dim UBoundFilmBlock As Long
    Dim c As Long
    UBoundFilmBlock = AllCurvesHeight \ FilmBlokSize
    ReDim FilmBlocks(UBoundFilmBlock)
    For q = 0 To UBoundFilmBlock
        Set FilmBlocks(q).Curves = New Collection
    Next
    For q = 1 To MyCurves.count
        Set MySortingCurve = MyCurves(q)
        c = MySortingCurve.path.Bounds.bottom \ FilmBlokSize
        If c > UBoundFilmBlock Then c = UBoundFilmBlock
        FilmBlocks(c).Curves.Add MySortingCurve
    Next

'    'Step 5
    Dim PreviosPoint As POINTL
    Dim LastAngel As Double
    Dim rc As New Collection
    Dim tempCol As Collection
    
    If lZeroPosition = RightP Then
        PreviosPoint.x = AllCurvesWidth
        LastAngel = 4.712389
    Else
        LastAngel = 1.5707965
    End If
    
    For q = 0 To UBoundFilmBlock
        If FilmBlocks(q).Curves.count > 0 Then
            Set tempCol = FindWay(FilmBlocks(q).Curves, PreviosPoint)
            Set MySortingCurve = tempCol(tempCol.count)
            PreviosPoint = MySortingCurve.path.LastPoint
            AddCollectionToCollection rc, tempCol
        End If
    Next
    Set Curves = SetCurves(rc)
    
    'Step 6
    CornerStep LastAngel
End Sub
