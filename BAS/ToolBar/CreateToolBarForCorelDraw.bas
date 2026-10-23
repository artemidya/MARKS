Private Sub CreateToolBarForCorelDraw()
    Select Case Application.VersionMajor
        Case 24 'CorelDraw 2022-2023
            CreateToolBarForCorelDrawX8
            
        Case 23 'CorelDraw 2021
            CreateToolBarForCorelDrawX8
            
        Case 22 'CorelDraw 2020
            CreateToolBarForCorelDrawX8
            
        Case 21 'CorelDraw 2019
            CreateToolBarForCorelDrawX8
            
        Case 20 'CorelDraw 2018
            CreateToolBarForCorelDrawX8
            
        Case 19 'CorelDraw 2017
            CreateToolBarForCorelDrawX8
            
        Case 18 'CorelDraw X8
            CreateToolBarForCorelDrawX8
            
        Case 17 'CorelDraw X7
            CreateToolBarForCorelDrawX7
            
        Case Else 'Corel ?
            CreateToolBarForCorelDrawX8
            
    End Select
End Sub
