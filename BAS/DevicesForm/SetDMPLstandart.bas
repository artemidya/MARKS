Private Sub SetDMPLstandart()
    Initialized = False
    Profiles(PlotterForm.TabStrip1.SelectedItem.index).Device = localization.ActiveLng.Item("Profile.DeviceDMPL")
    Profiles(PlotterForm.TabStrip1.SelectedItem.index).CommandLanguage = DMPL
    Text4.text = Profiles(PlotterForm.TabStrip1.SelectedItem.index).Device
    Combo2.ListIndex = ListIndexByName("DMPL", Combo2)
    ComboBox1.ListIndex = 0
    ComboBox6.ListIndex = 0
    CheckBox5.value = False
    TextBox15.text = ""
    TextBox16.text = ""
    CheckBox3.value = True
'    CheckBox4.value = False
    TextBox13.text = ""
    TextBox14.text = ""
    ComboBox3.ListIndex = 1
    TextBox1.text = "100"
    TextBox2.text = "100"
    CheckBox1.value = False
    CheckBox2.value = False
    ComboBox4.ListIndex = 0
    ComboBox5.ListIndex = 1
    TextBox3.text = "*.PLT"
    TextBox4.text = "4000000"
    TextBox5.text = PlotterLanguage.Init
    TextBox6.text = ""
    TextBox7.text = ""
    TextBox8.text = PlotterLanguage.TERM
    TextBox9.text = ""
    TextBox10.text = ""
    TextBox11.text = ""
    TextBox17.text = ""
    CheckBox6.value = False
    CheckBox7.value = True
    CheckBox7.Enabled = False
    
    TextBox1_AfterUpdate
    TextBox2_AfterUpdate
    TextBox4_AfterUpdate
    TextBox13_AfterUpdate
    TextBox14_AfterUpdate
    TextBox15_AfterUpdate
    TextBox16_AfterUpdate
    Initialized = True
End Sub
