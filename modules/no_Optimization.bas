Attribute VB_Name = "no_Optimization"
Sub noOpt() 'desabilitar opções para otimizar a execução do código
    Application.Calculation = xlCalculationAutomatic
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    Application.EnableEvents = True
End Sub
