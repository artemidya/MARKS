Private Sub UserForm_Initialize()
    ' Список распространённых скоростей передачи данных
    Dim baudRates As Variant
    baudRates = Array(110, 300, 1200, 2400, 4800, 9600, _
                      19200, 38400, 57600, 115200, 230400, 460800, 921600)
    
    Dim i As Integer
    For i = LBound(baudRates) To UBound(baudRates)
        Combo1.AddItem baudRates(i)
    Next i

    ComboBox1.AddItem 5
    ComboBox1.AddItem 6
    ComboBox1.AddItem 7
    ComboBox1.AddItem 8
    
    ComboBox2.AddItem "None"
    ComboBox2.AddItem "Odd"
    ComboBox2.AddItem "Even"
    ComboBox2.AddItem "Mark"
    ComboBox2.AddItem "Space"
    
    ComboBox3.AddItem "1"
    ComboBox3.AddItem "1.5"
    ComboBox3.AddItem "2"
    
    ComboBox4.AddItem "None"
    ComboBox4.AddItem "Xon/Xoff"
    ComboBox4.AddItem "RTS/CTS"
    ComboBox4.AddItem "DSR/DTR"
    
    
   
End Sub
