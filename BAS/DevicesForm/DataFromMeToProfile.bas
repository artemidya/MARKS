Private Sub DataFromMeToProfile()
    With Profiles(PlotterForm.TabStrip1.SelectedItem.index)
        On Error GoTo 10
        .Device = Text4.text
        .CommandLanguage = StringToCommandLanguage(Combo2.text)
        .ZeroPosition = ComboBox1.ListIndex
        .CornerPoints = CheckBox3.value
        .MaxSheetWidth = CLng(ValP(TextBox13.text))
        .MaxSheetHeight = CLng(ValP(TextBox14.text))
        .Turn = ComboBox3.ListIndex

        Dim sYorX As Double
        sYorX = ValP(TextBox1.text) / 100
        If CheckBox1.value Then sYorX = -sYorX
        .HeightScale = sYorX

        sYorX = ValP(TextBox2.text) / 100
        If CheckBox2.value Then sYorX = -sYorX
        .WidthScale = sYorX

        .LineType = ComboBox4.ListIndex
        .ApproximationLevel = ComboBox5.ListIndex
        .FilePLT = TextBox3.text

        .MemorySize = CLng(val(TextBox4.text))
        .Initializer = TextBox5.text
        .BeginWorkFile = TextBox6.text
        .EndWorkFile = TextBox7.text
        .Terminator = TextBox8.text
        
        .AccelerationsAsString = TextBox9.text
        .SpeedsAsString = TextBox10.text
        .ForcesAsString = TextBox11.text
        .Laser = CheckBox5.value
        .LaserDX = CLng(val(TextBox15.text))
        .LaserDY = CLng(val(TextBox16.text))
        .SeparateParametrs = CheckBox6.value
        .OneCommandOnPath = CheckBox7.value
        .PlotterModelString = TextBox17.text

        If ComboBox6.ListIndex = 1 Then
            .CoordinateSystem = Relative
        Else
            .CoordinateSystem = Absolute
        End If
10
    End With
End Sub
