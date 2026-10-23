Private Sub Command5_Click()
    dntChange = True
    ' Установка значения по умолчанию
    Combo1.value = 9600
    ComboBox1.value = 8
    ComboBox2.value = "None"
    ComboBox3.value = "1"
    ComboBox4.value = "None"
    dntChange = False
    BuildComSettingsString
    CheckDefaultButton
End Sub
