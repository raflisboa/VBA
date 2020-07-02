VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Consulta
   Caption         =   "Consultar Dados - Logbook de An�lises"
   ClientHeight    =   9150.001
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   14970
   OleObjectBlob   =   "Consulta.frx":0000
   StartUpPosition =   2  'CenterScreen
End
Attribute VB_Name = "Consulta"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public MatrizResultados As Variant
Public Total_Ocorrencias As Long

' rotina para mostrar a imagem ou local de imagem armazenado na an�lise

Private Sub botao_link_Click()
Call Opt

  Link = caixa_link.Text()

  ' condi��o que impede que a pasta de fotos seja aberta sem nenhuma imagem associada no formul�rio
  If Link = "\\Brpcesfil01\DEBUG_EN\1 - Debug\4 - Documentos Gerais\3 - Banco de Arquivos Compartilhados\2 - Imagens\" Then

    GoTo NoCanDo

     End If

   On Error GoTo NoCanDo

   ActiveWorkbook.FollowHyperlink Address:=caixa_link.Object, ExtraInfo:="ID=0", Method:=msoMethodPost, NewWindow:=False

   Exit Sub

NoCanDo:

    MsgBox "N�o h� imagem associada a essa an�lise ", vbCritical

Call noOpt

End Sub

Private Sub botao_pesquisaModelo_Click()

Call Opt


    If Me.combo_modelo.Text = "" Then
        MsgBox "Escolha um modelo para pesquisar"

    Else
        Call ProcuraPersonalizada(Me.combo_modelo.Text)
    End If

Call noOpt


End Sub

Private Sub botao_procurar_Click()

Call Opt


    If Me.txt_Procurar.Text = "" Then
        MsgBox "Digite um termo para pesquisar"
   ' ElseIf Me.combo_modelo.Text = "" Then

    Else
        Call ProcuraPersonalizada(Me.txt_Procurar.Text)
    End If

Call noOpt

End Sub

Private Sub caixa_resultadoPPID_Change()

End Sub

Private Sub caixa_resultadoModelo_Change()

End Sub

Private Sub caixa_resultadoSemana_Change()

End Sub

Private Sub caixa_resultadoEstacaoFalha_Change()

End Sub

Private Sub caixa_pesquisa_Change()

End Sub

Private Sub CommandButton1_Click()

End Sub

Private Sub botao_sair_Click() ' rotina para sair da tela de consulta com seguran�a e voltar a tela de login

    Application.DisplayAlerts = False

        Sheet2.Visible = xlSheetVisible
        Plan2.Visible = xlSheetVeryHidden
        Plan1.Visible = xlSheetVeryHidden

            Unload Me

    UserLogin.Show

End Sub

Private Sub botao_tecnico_Click()
Call Opt

    If Me.combo_tecnico.Text = "" Then
        MsgBox "Escolha um t�cnico para pesquisar"

    Else
        Call ProcuraPersonalizada(Me.combo_tecnico.Text)
    End If


Call noOpt

End Sub

Private Sub caixa_estacao_Change()

End Sub

Private Sub caixa_link_Change()

End Sub

Private Sub caixa_outras_Change()

End Sub

Private Sub caixa_outrosComponentes_Change()

End Sub

Private Sub caixa_posicaoComponente_Change()

End Sub

Private Sub caixa_resultadoTipoFalha_Change()

End Sub

Private Sub caixa_sintomas_Change()

End Sub

Private Sub caixa_tecnico_Change()

End Sub

Private Sub ComboBox10_Change()

End Sub

Private Sub ComboBox12_Change()

End Sub

Private Sub combo_modelo_Change()

End Sub

Private Sub combo_tecnico_Change()

End Sub

Private Sub digite_Click()

End Sub

Private Sub Image1_Click()

End Sub

Private Sub Label1_Click()

End Sub

Private Sub Label14_Click()

End Sub


Private Sub Label15_Click()

End Sub

Private Sub Label9_Click()

End Sub

Private Sub SpinButton1_Change() ' propriedades do bot�o de navega��o pelos resultados da consulta

Call Opt

Dim Linha As Long
Dim TotalOcorrencias As Long


    TotalOcorrencias = SpinButton1.Max + 1

        Linha = MatrizResultados(SpinButton1.Value)
        Label_Registros_Contador.Caption = SpinButton1.Value + 1 & " de " & TotalOcorrencias

            caixa_ppid.Text = Plan1.Cells(Linha, 1).Value
            caixa_modelo.Text = Plan1.Cells(Linha, 2).Value
            caixa_semana.Text = Plan1.Cells(Linha, 3).Value
            caixa_estacao.Text = Plan1.Cells(Linha, 4).Value
            caixa_resultadoTipoFalha.Text = Plan1.Cells(Linha, 5).Value
            caixa_sintomas.Text = Plan1.Cells(Linha, 6).Value
            caixa_sinais.Text = Plan1.Cells(Linha, 7).Value
            caixa_posicaoComponente.Text = Plan1.Cells(Linha, 8).Value
            caixa_tipoComponente.Text = Plan1.Cells(Linha, 9).Value
            caixa_tipoReparo.Text = Plan1.Cells(Linha, 10).Value
            caixa_outras.Text = Plan1.Cells(Linha, 11).Value
            caixa_tecnico.Text = Plan1.Cells(Linha, 12).Value
            caixa_link.Text = Plan1.Cells(Linha, 13).Value
            caixa_outrosComponentes.Text = Plan1.Cells(Linha, 14).Value

Call noOpt

End Sub


Private Sub ProcuraPersonalizada(ByVal TermoPesquisado As String)

Call Opt


Dim Busca As Range
Dim Primeira_Ocorrencia As String
Dim Resultados As String

    'Executa a busca
    Set Busca = Plan1.Cells.Find(What:=TermoPesquisado, After:=Range("A1"), LookIn:=xlFormulas, _
        LookAt:=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, _
        MatchCase:=False, SearchFormat:=False)

    'Caso tenha encontrado alguma ocorr�ncia...
    If Not Busca Is Nothing Then

        Primeira_Ocorrencia = Busca.Address
        Resultados = Busca.Row  'Lista o primeiro resultado na variavel

        'Neste loop, pesquisa todas as pr�ximas ocorr�ncias para
        'o termo pesquisado
        Do
            Set Busca = Plan1.Cells.FindNext(After:=Busca)

            'Condicional para n�o listar o primeiro resultado
            'pois j� foi listado acima
            If Not Busca.Address Like Primeira_Ocorrencia Then
                Resultados = Resultados & ";" & Busca.Row
            End If
        Loop Until Busca.Address Like Primeira_Ocorrencia


        MatrizResultados = Split(Resultados, ";")

        'Atualiza dados iniciais no formul�rio
            SpinButton1.Max = UBound(MatrizResultados)  'Valor maximo do seletor de registros

        'habilita o seletor de registro
                    SpinButton1.Enabled = True

        'indicador do seletor de registros
                        Label_Registros_Contador.Caption = "1 de " & UBound(MatrizResultados) + 1


        'Box com o conteudo encontrado
        caixa_ppid.Text = Plan1.Cells(MatrizResultados(0), 1).Value
        caixa_modelo.Text = Plan1.Cells(MatrizResultados(0), 2).Value
        caixa_semana.Text = Plan1.Cells(MatrizResultados(0), 3).Value
        caixa_estacao.Text = Plan1.Cells(MatrizResultados(0), 4).Value
        caixa_resultadoTipoFalha.Text = Plan1.Cells(MatrizResultados(0), 5).Value
        caixa_sintomas.Text = Plan1.Cells(MatrizResultados(0), 6).Value
        caixa_sinais.Text = Plan1.Cells(MatrizResultados(0), 7).Value
        caixa_posicaoComponente.Text = Plan1.Cells(MatrizResultados(0), 8).Value
        caixa_tipoComponente.Text = Plan1.Cells(MatrizResultados(0), 9).Value
        caixa_tipoReparo.Text = Plan1.Cells(MatrizResultados(0), 10).Value
        caixa_outras.Text = Plan1.Cells(MatrizResultados(0), 11).Value
        caixa_tecnico.Text = Plan1.Cells(MatrizResultados(0), 12).Value
        caixa_tecnico.Text = Plan1.Cells(MatrizResultados(0), 12).Value
        caixa_link.Text = Plan1.Cells(MatrizResultados(0), 13).Value
        caixa_outrosComponentes.Text = Plan1.Cells(MatrizResultados(0), 14).Value


    Else    'Caso nada tenha sido encontrado, exibe mensagem informativa

        SpinButton1.Enabled = False     'desabilita o seletor de registros

            Label_Registros_Contador.Caption = ""   'zera os resultados encontrados

                'limpa os campos do formul�rio
                caixa_ppid.Text = ""
                caixa_modelo.Text = ""
                caixa_semana.Text = ""
                caixa_estacao.Text = ""
                caixa_resultadoTipoFalha.Text = ""
                caixa_sintomas.Text = ""
                caixa_sinais.Text = ""
                caixa_posicaoComponente.Text = ""
                caixa_tipoComponente.Text = ""
                caixa_tipoReparo.Text = ""
                caixa_outras.Text = ""
                caixa_tecnico.Text = ""
                caixa_outrosComponentes.Text = ""

        MsgBox "Nenhum resultado para '" & TermoPesquisado & "' foi encontrado."

    End If

Call noOpt

End Sub

Private Sub caixa_ppid_Change()

End Sub

Private Sub txt_Procurar_Change()

End Sub

Private Sub UserForm_Initialize()

Call Apresentar_on


Call Opt

' preenchendo as Combo Box - Dinamicamente

lin = 2
    Do Until Sheets("Aux_1").Cells(lin, 6) = ""
    combo_modelo.AddItem Sheets("Aux_1").Cells(lin, 6)
    lin = lin + 1

    Loop

lin = 2

    Do Until Sheets("Aux_1").Cells(lin, 4) = ""
    combo_tecnico.AddItem Sheets("Aux_1").Cells(lin, 4)
    lin = lin + 1

    Loop


'rotinas para navegar entre os resultados

    SpinButton1.Enabled = False

    Label_Registros_Contador.Caption = ""

Call noOpt

End Sub


'cancela o bot�o de fechamento do Userform
Private Sub UserForm_QueryClose _
  (Cancel As Integer, CloseMode As Integer)

Call Opt

    If CloseMode = vbFormControlMenu Then
        Cancel = True
    End If

Call noOpt

End Sub
