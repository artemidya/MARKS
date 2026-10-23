Public Sub RefreshFields()
    If TabStrip1.SelectedItem.index = 8 Then Exit Sub
    
    With Profiles(TabStrip1.SelectedItem.index)
        Text4.text = .Device
        Combo1.ListIndex = ListIndexByName(.OutPut, Combo1)
        TextBox9.text = Format((.ClosingCurve / 1000), "0.000") & " " & localization.ActiveLng.Item("RaskroyForm.mm")
        TextBox10.text = Format((.KnifeOffset / 1000), "0.000") & " " & localization.ActiveLng.Item("RaskroyForm.mm")
        CheckBox9.value = RegData1.PreRun
        CheckBox10.value = .DoublePass
        Dim q As Long
        '====Accelerations====
        With ComboBox1
            .Clear
            .AddItem localization.ActiveLng.Item("PlotterForm.ComboBox1")
            For q = 0 To Profiles(TabStrip1.SelectedItem.index).AccelerationsCount - 1
                .AddItem Replace(Profiles(TabStrip1.SelectedItem.index).Accelerations(q) / 10 & " " & localization.ActiveLng.Item("PlotterForm.m/s") & ChrW(178), ",", ".")
            Next
            
            If .ListCount = 1 Or (Profiles(TabStrip1.SelectedItem.index).CommandLanguage = DMPL) Then
                .Enabled = False
                .ListIndex = 0
                Label14.Enabled = False
                ImageAcceleration.Picture = ImageAccelerationDisabled.Picture
            Else
                .ListIndex = Profiles(TabStrip1.SelectedItem.index).ActiveAcceleration
                .Enabled = True
                Label14.Enabled = True
                ImageAcceleration.Picture = ImageAccelerationEnabled.Picture
            End If
        End With
        '====Speeds====
        With ComboBox2
            If Profiles(TabStrip1.SelectedItem.index).SeparateParametrs Then
                .Visible = False
                ComboBox6.Visible = True
                ComboBox7.Visible = True
            Else
                .Visible = True
                ComboBox6.Visible = False
                ComboBox7.Visible = False
            End If
            
            .Clear
            ComboBox6.Clear
            ComboBox7.Clear
            
            .AddItem localization.ActiveLng.Item("PlotterForm.ComboBox1")
            ComboBox6.AddItem localization.ActiveLng.Item("PlotterForm.ComboBox1")
            ComboBox7.AddItem localization.ActiveLng.Item("PlotterForm.ComboBox1")
            
            For q = 0 To Profiles(TabStrip1.SelectedItem.index).SpeedsCount - 1
                .AddItem Profiles(TabStrip1.SelectedItem.index).Speeds(q) & " " & localization.ActiveLng.Item("PlotterForm.mm/s")
                ComboBox6.AddItem ChrW(9660) & " " & Profiles(TabStrip1.SelectedItem.index).Speeds(q) & " " & localization.ActiveLng.Item("PlotterForm.mm/s")
                ComboBox7.AddItem ChrW(9650) & " " & Profiles(TabStrip1.SelectedItem.index).Speeds(q) & " " & localization.ActiveLng.Item("PlotterForm.mm/s")
            Next
            If .ListCount = 1 Then
                .ListIndex = 0
                ComboBox6.ListIndex = 0
                ComboBox7.ListIndex = 0
                .Enabled = False
                ComboBox6.Enabled = False
                ComboBox7.Enabled = False
                Label12.Enabled = False
                ImageSpeed.Picture = ImageSpeedDisebled.Picture
            Else
                .ListIndex = Profiles(TabStrip1.SelectedItem.index).ActiveSpeed
                ComboBox6.ListIndex = Profiles(TabStrip1.SelectedItem.index).ActiveSpeed
                ComboBox7.ListIndex = Profiles(TabStrip1.SelectedItem.index).ActiveSpeedUP
                
                .Enabled = True
                ComboBox6.Enabled = True
                ComboBox7.Enabled = True
                Label12.Enabled = True

                ImageSpeed.Picture = ImageSpeedEnabled.Picture
            End If
        End With
        
        '====Forces====
        With ComboBox3
            .Clear
            .AddItem localization.ActiveLng.Item("PlotterForm.ComboBox1")
            For q = 0 To Profiles(TabStrip1.SelectedItem.index).ForcesCount - 1
                .AddItem Profiles(TabStrip1.SelectedItem.index).Forces(q) & " " & localization.ActiveLng.Item("PlotterForm.g")
            Next
            If .ListCount = 1 Then
                .Enabled = False
                .ListIndex = 0
                Label13.Enabled = False
                ImageForce.Picture = ImageForceDisabled.Picture
            Else
                .ListIndex = ListIndexByName(Profiles(TabStrip1.SelectedItem.index).ActiveForce & " " & localization.ActiveLng.Item("PlotterForm.g"), ComboBox3)
                .Enabled = True
                Label13.Enabled = True
                ImageForce.Picture = ImageForceEnabled.Picture
            End If
        End With
        If ComboBox1.Enabled = False And ComboBox2.Enabled = False And ComboBox3.Enabled = False Then
            Frame4.Enabled = False
        Else
            Frame4.Enabled = True
        End If
    End With
    RefreshXY
End Sub
