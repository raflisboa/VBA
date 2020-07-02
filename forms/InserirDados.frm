VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} InserirDados 
   Caption         =   "Inserir Dados - Logbook de Análises"
   ClientHeight    =   10215
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   10350
   OleObjectBlob   =   "InserirDados.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "InserirDados"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub botao_cancelar_Click()

    Application.DisplayAlerts = False
    
    'ActiveWorkbook.Save
    
    Unload Me
    
    UserLogin.Show
End Sub
Private Sub botao_limpar_Click()

    Call UserForm_Initialize
    
End Sub

Private Sub sintomas_Click()

End Sub

Private Sub caixa_linkImagem_Change()

End Sub

Private Sub caixa_outrosComponentes_Change()

End Sub

Private Sub caixa_semana_Change()

End Sub

Private Sub caixa_sintomas_Change()

End Sub

Private Sub combo_estacao_Change()

End Sub

Private Sub combo_tipoComponente_Change()

End Sub

Private Sub componentes_Click()

End Sub

Private Sub Image1_Click()

End Sub

Private Sub Label8_Click()

End Sub

Private Sub caixa_modelo_Change()

End Sub

Private Sub caixa_observacoes_Change()

End Sub

Private Sub caixa_ppid_Change()

End Sub

Private Sub combo_modelo_Change()

End Sub

Private Sub combo_tecnico_Change()

End Sub

Private Sub ComboBox1_Change()

End Sub

Private Sub combo_tipo_Change()

End Sub

Private Sub Label5_Click()

End Sub

Private Sub TextBox1_Change()

End Sub

    
Private Sub UserForm_Initialize()
    
Call Opt


'Retorna os campos ao Default
    caixa_ppid.Value = ""
    combo_modelo.Value = ""
    caixa_semana.Value = ""
    caixa_sintomas.Value = ""
    caixa_sinais.Value = ""
    caixa_componentes.Value = ""
    caixa_observacoes.Value = ""
    combo_tipoReparo.Value = ""
    caixa_linkImagem.Value = ""
    combo_tipoComponente.Value = ""
    combo_tipo.Value = ""
    combo_estacao.Value = ""
    caixa_outrosComponentes.Value = ""

' preenchendo as Combo Box - Dinamicamente


lin = 2
    Do Until Sheets("Aux_1").Cells(lin, 1) = ""
    combo_estacao.AddItem Sheets("Aux_1").Cells(lin, 1)
    lin = lin + 1

    Loop

lin = 2
    Do Until Sheets("Aux_1").Cells(lin, 2) = ""
    combo_tipo.AddItem Sheets("Aux_1").Cells(lin, 2)
    lin = lin + 1

    Loop

lin = 2
    Do Until Sheets("Aux_1").Cells(lin, 3) = ""
    combo_tipoReparo.AddItem Sheets("Aux_1").Cells(lin, 3)
    lin = lin + 1

    Loop

lin = 2
    Do Until Sheets("Aux_1").Cells(lin, 4) = ""
    combo_tecnico.AddItem Sheets("Aux_1").Cells(lin, 4)
    lin = lin + 1

    Loop

lin = 2
    Do Until Sheets("Aux_1").Cells(lin, 5) = ""
    combo_tipoComponente.AddItem Sheets("Aux_1").Cells(lin, 5)
    lin = lin + 1

    Loop

lin = 2
    Do Until Sheets("Aux_1").Cells(lin, 6) = ""
    combo_modelo.AddItem Sheets("Aux_1").Cells(lin, 6)
    lin = lin + 1

    Loop


'Configuração para manter o cursor de edição ativo no campo PPID
        caixa_ppid.SetFocus

            Call noOpt


End Sub

Private Sub botao_enviar_Click()
Call Opt


'Cria a variavel linhavazia
Dim linhavazia As Long



'Confere se o campo PPID foi preenchido

If caixa_ppid.Value = "" Then
    MsgBox ("Campo PPID é obrigatório")
        caixa_ppid.SetFocus
        Exit Sub
    Else
End If

'seleciona a aba "dados"
'Plan1.Visible = xlSheetVisible

Plan1.Activate


'conta quantas informações foram inseridas na coluna A da aba dados
linhavazia = WorksheetFunction.CountA(Range("A:A")) + 1

'Insere informações da aba dados
        Cells(linhavazia, 1).Value = caixa_ppid.Value
        Cells(linhavazia, 2).Value = combo_modelo.Value
        Cells(linhavazia, 3).Value = caixa_semana.Value & "/2017"
        Cells(linhavazia, 4).Value = combo_estacao.Value
        Cells(linhavazia, 5).Value = combo_tipo.Value
        Cells(linhavazia, 6).Value = caixa_sintomas.Value
        Cells(linhavazia, 7).Value = caixa_sinais.Value
        Cells(linhavazia, 8).Value = caixa_componentes.Value
        Cells(linhavazia, 9).Value = combo_tipoReparo.Value
        Cells(linhavazia, 10).Value = caixa_observacoes.Value
        Cells(linhavazia, 11).Value = combo_tecnico.Value
        Cells(linhavazia, 12).Value = combo_tipoComponente.Value
        Cells(linhavazia, 13).Value = "\\Brpcesfil01\DEBUG_EN\1 - Debug\4 - Documentos Gerais\3 - Banco de Arquivos Compartilhados\2 - Imagens\" & caixa_linkImagem.Value
        Cells(linhavazia, 14).Value = caixa_outrosComponentes.Value


'Avisa que informações foi inserida com sucesso
MsgBox ("Informação inserida com sucesso")

'Volta para a aba MENU
'Worksheets("MENU").Select

'Retorna os  campos ao defaut
        caixa_ppid.Value = ""
        combo_modelo.Value = ""
        caixa_semana.Value = ""
        caixa_sintomas.Value = ""
        caixa_sinais.Value = ""
        caixa_componentes.Value = ""
        caixa_observacoes.Value = ""
        combo_tipoReparo.Value = ""
        combo_tipoComponente.Value = ""
        combo_tipo.Value = ""
        combo_tecnico.Value = ""
        combo_estacao.Value = ""
        caixa_outrosComponentes.Value = ""
        caixa_linkImagem.Value = ""

    Application.DisplayAlerts = False
    ActiveWorkbook.Save

' esconde a aba "Banco de dados"
        Plan1.Visible = xlSheetVeryHidden

            Call noOpt


End Sub

'cancela o botão de fechamento do Userform

Private Sub UserForm_QueryClose _
  (Cancel As Integer, CloseMode As Integer)
Call Opt

    If CloseMode = vbFormControlMenu Then
        Cancel = True
    End If
    
Call noOpt
 
End Sub

