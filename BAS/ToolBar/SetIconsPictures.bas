Private Sub SetIconsPictures()
    If Dir(Application.GMSManager.Projects("PlotCalc").FilePath & "icons", vbDirectory) = "" Then MkDir Application.GMSManager.Projects("PlotCalc").FilePath & "icons"
    If Dir(Application.GMSManager.Projects("PlotCalc").FilePath & "icons\14.png") = "" Then CreateFile_14_png
    If Dir(Application.GMSManager.Projects("PlotCalc").FilePath & "icons\0.png") = "" Then CreateFile_0_png
    If Dir(Application.GMSManager.Projects("PlotCalc").FilePath & "icons\0_5.png") = "" Then CreateFile_0_5_png
    If Dir(Application.GMSManager.Projects("PlotCalc").FilePath & "icons\1.png") = "" Then CreateFile_1_png
    If Dir(Application.GMSManager.Projects("PlotCalc").FilePath & "icons\2.png") = "" Then CreateFile_2_png
    If Dir(Application.GMSManager.Projects("PlotCalc").FilePath & "icons\3.png") = "" Then CreateFile_3_png
    If Dir(Application.GMSManager.Projects("PlotCalc").FilePath & "icons\4.png") = "" Then CreateFile_4_png
    If Dir(Application.GMSManager.Projects("PlotCalc").FilePath & "icons\5.png") = "" Then CreateFile_5_png
    If Dir(Application.GMSManager.Projects("PlotCalc").FilePath & "icons\6.png") = "" Then CreateFile_6_png
    If Dir(Application.GMSManager.Projects("PlotCalc").FilePath & "icons\7.png") = "" Then CreateFile_7_png
    If Dir(Application.GMSManager.Projects("PlotCalc").FilePath & "icons\9.png") = "" Then CreateFile_9_png
    If Dir(Application.GMSManager.Projects("PlotCalc").FilePath & "icons\10.png") = "" Then CreateFile_10_png
    If Dir(Application.GMSManager.Projects("PlotCalc").FilePath & "icons\11.png") = "" Then CreateFile_11_png
    If Dir(Application.GMSManager.Projects("PlotCalc").FilePath & "icons\8.png") = "" Then CreateFile_8_png
    If Dir(Application.GMSManager.Projects("PlotCalc").FilePath & "icons\12.png") = "" Then CreateFile_12_png
    If Dir(Application.GMSManager.Projects("PlotCalc").FilePath & "icons\13.png") = "" Then CreateFile_13_png
    
    Application.CommandBars("PlotCalc").Controls(1).SetIcon2 Application.GMSManager.Projects("PlotCalc").FilePath & "icons\14.png"
    Application.CommandBars("PlotCalc").Controls(2).SetIcon2 Application.GMSManager.Projects("PlotCalc").FilePath & "icons\0.png"
    Application.CommandBars("PlotCalc").Controls(3).SetIcon2 Application.GMSManager.Projects("PlotCalc").FilePath & "icons\0_5.png"
    Application.CommandBars("PlotCalc").Controls(4).SetIcon2 Application.GMSManager.Projects("PlotCalc").FilePath & "icons\1.png"
    Application.CommandBars("PlotCalc").Controls(5).SetIcon2 Application.GMSManager.Projects("PlotCalc").FilePath & "icons\2.png"
    Application.CommandBars("PlotCalc").Controls(6).SetIcon2 Application.GMSManager.Projects("PlotCalc").FilePath & "icons\3.png"
    Application.CommandBars("PlotCalc").Controls(7).SetIcon2 Application.GMSManager.Projects("PlotCalc").FilePath & "icons\4.png"
    Application.CommandBars("PlotCalc").Controls(8).SetIcon2 Application.GMSManager.Projects("PlotCalc").FilePath & "icons\5.png"
    Application.CommandBars("PlotCalc").Controls(9).SetIcon2 Application.GMSManager.Projects("PlotCalc").FilePath & "icons\6.png"
    Application.CommandBars("PlotCalc").Controls(10).SetIcon2 Application.GMSManager.Projects("PlotCalc").FilePath & "icons\7.png"
    Application.CommandBars("PlotCalc").Controls(11).SetIcon2 Application.GMSManager.Projects("PlotCalc").FilePath & "icons\9.png"
    Application.CommandBars("PlotCalc").Controls(12).SetIcon2 Application.GMSManager.Projects("PlotCalc").FilePath & "icons\10.png"
    Application.CommandBars("PlotCalc").Controls(13).SetIcon2 Application.GMSManager.Projects("PlotCalc").FilePath & "icons\13.png"
    Application.CommandBars("PlotCalc").Controls(14).SetIcon2 Application.GMSManager.Projects("PlotCalc").FilePath & "icons\12.png"
    Application.CommandBars("PlotCalc").Controls(15).SetIcon2 Application.GMSManager.Projects("PlotCalc").FilePath & "icons\11.png"
    Application.CommandBars("PlotCalc").Controls(16).SetIcon2 Application.GMSManager.Projects("PlotCalc").FilePath & "icons\8.png"
End Sub
