Attribute VB_Name = "ApresentarOn"
Sub Apresentar_on()
    Application.ExecuteExcel4Macro "SHOW.TOOLBAR(""Ribbon"",False)" 'Oculta todas as guias de menu
        Application.DisplayFormulaBar = False 'Ocultar barra de fórmulas
        Application.DisplayStatusBar = False 'Ocultar barra de status, disposta ao final da planilha
            ActiveWindow.DisplayHeadings = False 'Ocultar o cabeçalho da Pasta de trabalho
                
                With ActiveWindow
                    .DisplayWorkbookTabs = False 'Ocultar guias das planilhas
                    .DisplayHeadings = False 'Oculta os títulos de linha e coluna
                    .DisplayHorizontalScrollBar = False 'Ocultar barra horizontal
                    .DisplayVerticalScrollBar = False 'Ocultar barra vertical
                    .DisplayZeros = False 'Oculta valores zero na planilha
                    .DisplayGridlines = False 'Oculta as linhas de grade
                End With
End Sub
