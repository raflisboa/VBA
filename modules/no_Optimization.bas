Attribute VB_Name = "no_Optimization"
Sub noOpt() 'desabilitar op��es para otimizar a execu��o do c�digo
    Application.Calculation = xlCalculationAutomatic
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    Application.EnableEvents = True
End Sub
