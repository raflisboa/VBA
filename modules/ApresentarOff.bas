Attribute VB_Name = "Apresentaroff"
Sub Apresentar_off()
    Application.ExecuteExcel4Macro "SHOW.TOOLBAR(""Ribbon"",True)" 'Mostrar todas as guias de menu
        Application.DisplayFormulaBar = True 'Mostrar a barra de f�rmulas
        Application.DisplayStatusBar = True 'Mostrar a barra de status, disposta ao final da planilha
            ActiveWindow.DisplayHeadings = True 'Mostrar o cabe�alho da Pasta de trabalho
                
                With ActiveWindow
                    .DisplayWorkbookTabs = True 'Mostrar guias das planilhas
                    .DisplayHeadings = True 'Mostrar os t�tulos de linha e coluna
                    .DisplayHorizontalScrollBar = True 'Mostrar barra horizontal
                    .DisplayVerticalScrollBar = True 'Mostrar barra vertical
                    .DisplayZeros = True 'Mostrar valores zero na planilha
                End With
End Sub
