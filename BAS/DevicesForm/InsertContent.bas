Private Sub InsertContent(ByRef content As String)
    On Error Resume Next
    Initialized = False
    Text4.text = GetTag("Device", content)
    Combo2.ListIndex = ListIndexByName(GetTag("CommandLanguage", content), Combo2)
    ComboBox1.ListIndex = StringToLong(GetTag("ZeroPosition", content))
    CheckBox3.value = StringToBoolean(GetTag("CornerPoints", content))
    
    Dim sc As String
    sc = GetTag("CoordinateSystem", content)
    If UCase(sc) = "RELATIVE" Then
        ComboBox6.ListIndex = 1
    Else
        ComboBox6.ListIndex = 0
    End If
    
    TextBox13.text = StringToLong(GetTag("MaxSheetWidth", content))
    TextBox14.text = StringToLong(GetTag("MaxSheetHeight", content))
    ComboBox3.ListIndex = StringToLong(GetTag("Turn", content))
    
    Dim HeightScale As Double, WidthScale As Double
    HeightScale = ValPD(GetTag("HeightScale", content))
    WidthScale = ValPD(GetTag("WidthScale", content))
    TextBox1.text = Format(Abs(HeightScale * 100), "0.00") & "%"
    TextBox2.text = Format(Abs(WidthScale * 100), "0.00") & "%"
    If HeightScale < 0 Then CheckBox1.value = True Else CheckBox1.value = False
    If WidthScale < 0 Then CheckBox2.value = True Else CheckBox2.value = False
    
'    ComboBox4.ListIndex = StringToLong(GetTag("LineType", content))
    ComboBox4.ListIndex = 0
    ComboBox5.ListIndex = StringToLong(GetTag("ApproximationLevel", content))
    TextBox3.text = GetTag("FilePLT", content)
    TextBox4.text = StringToLong(GetTag("MemorySize", content))
    TextBox5.text = GetTag("Initializer", content)
    TextBox6.text = GetTag("BeginWorkFile", content)
    TextBox7.text = GetTag("EndWorkFile", content)
    TextBox8.text = GetTag("Terminator", content)
    
    Dim arrStr() As String
    arrStr = Split(GetTag("Accelerations", content), ";")
    TextBox9.text = ""
    If UBound(arrStr) > -1 Then
        Dim q As Long
        For q = 0 To UBound(arrStr) - 1
            TextBox9.text = TextBox9.text & Replace(StringToLong(arrStr(q)) / 10, ",", ".") & ";"
        Next
        TextBox9.text = TextBox9.text & Replace(StringToLong(arrStr(q)) / 10, ",", ".")
    End If
    
    TextBox10.text = GetTag("Speeds", content)
    TextBox11.text = GetTag("Forces", content)
    CheckBox5.value = StringToBoolean(GetTag("Laser", content))
    TextBox15.text = StringToLong(GetTag("LaserDX", content))
    TextBox16.text = StringToLong(GetTag("LaserDY", content))
    CheckBox6.value = StringToBoolean(GetTag("SeparateParametrs", content))
    If Combo2.ListIndex = 1 Then
        CheckBox7.Enabled = False
        CheckBox7.value = True
    Else
        CheckBox7.Enabled = True
        CheckBox7.value = StringToBoolean(GetTag("OneCommandOnPath", content))
    End If
    TextBox13_AfterUpdate
    TextBox14_AfterUpdate
    TextBox15_AfterUpdate
    TextBox16_AfterUpdate
    TextBox17.text = GetTag("PlotterModelString", content)
    Initialized = True
End Sub
