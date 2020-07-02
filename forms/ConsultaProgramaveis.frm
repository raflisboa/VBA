VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ConsultaProgramaveis 
   Caption         =   "Double Check Programaveis - BIOS ROOM - nPCEBG - Jundiai Site"
   ClientHeight    =   6435
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   12075
   OleObjectBlob   =   "ConsultaProgramaveis.frx":0000
   StartUpPosition =   2  'CenterScreen
End
Attribute VB_Name = "ConsultaProgramaveis"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False



Public MatrizResultados As Variant
Public Total_Ocorrencias As Long

' rotina para mostrar a imagem ou local de imagem armazenado na análise

Private Sub botao_link_Click()
Call Opt

  Link = caixa_link.Text()
  
  ' condição que impede que a pasta de fotos seja aberta sem nenhuma imagem associada no formulário
  If Link = "\\Brpcesfil01\DEBUG_EN\1 - Debug\3 - OSV+\3 - Banco de Arquivos\1 - Imagens" Then
   
    GoTo NoCanDo
    
     End If
    
   On Error GoTo NoCanDo
    
   ActiveWorkbook.FollowHyperlink Address:=caixa_link.Object, ExtraInfo:="ID=0", Method:=msoMethodPost, NewWindow:=False
        
   Exit Sub
   
NoCanDo:

    MsgBox "Não há imagem associada a essa análise ", vbCritical
    
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

Private Sub botao_sair_Click() ' rotina para sair da tela de consulta com segurança e voltar a tela de login

    Application.DisplayAlerts = False
    
        Sheet2.Visible = xlSheetVisible
        Plan2.Visible = xlSheetVeryHidden
        Sheet3.Visible = xlSheetVeryHidden

            Unload Me

    UserLogin.Show

End Sub

Private Sub botao_tecnico_Click()
Call Opt

    If Me.combo_tecnico.Text = "" Then
        MsgBox "Escolha um técnico para pesquisar"
         
    Else
        Call ProcuraPersonalizada(Me.combo_tecnico.Text)
    End If

   
Call noOpt

End Sub

Private Sub caixa_estacao_Change()

End Sub

Private Sub caixa_link_Change()

End Sub

Private Sub caixa_modelo_Change()

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

Private Sub Image2_BeforeDragOver(ByVal Cancel As MSForms.ReturnBoolean, ByVal Data As MSForms.DataObject, ByVal X As Single, ByVal Y As Single, ByVal DragState As MSForms.fmDragState, ByVal Effect As MSForms.ReturnEffect, ByVal Shift As Integer)

End Sub

Private Sub Image2_Click()

    Application.DisplayAlerts = False
    
        Sheet2.Visible = xlSheetVisible
        Plan2.Visible = xlSheetVeryHidden
        Sheet3.Visible = xlSheetVeryHidden

            Unload Me

    UserLogin.Show


End Sub

Private Sub Label14_Click()

End Sub


Private Sub Label15_Click()

End Sub

Private Sub Label18_Click()

End Sub

Private Sub Label2_Click()

End Sub

Private Sub Label9_Click()

End Sub

Private Sub SpinButton1_Change() ' propriedades do botão de navegação pelos resultados da consulta

Call Opt

Dim Linha As Long
Dim TotalOcorrencias As Long


    TotalOcorrencias = SpinButton1.Max + 1
    
        Linha = MatrizResultados(SpinButton1.Value)
        Label_Registros_Contador.Caption = SpinButton1.Value + 1 & " de " & TotalOcorrencias
            
            caixa_ppid.Text = Sheet3.Cells(Linha, 1).Value
            caixa_modelo.Text = Sheet3.Cells(Linha, 2).Value
            caixa_link.Text = Sheet3.Cells(Linha, 3).Value
        
Call noOpt
    
End Sub


Private Sub ProcuraPersonalizada(ByVal TermoPesquisado As String)

Call Opt


Dim Busca As Range
Dim Primeira_Ocorrencia As String
Dim Resultados As String

    'Executa a busca
    Set Busca = Sheet3.Cells.Find(What:=TermoPesquisado, After:=Range("A1"), LookIn:=xlFormulas, _
        LookAt:=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, _
        MatchCase:=False, SearchFormat:=False)
    
    'Caso tenha encontrado alguma ocorrência...
    If Not Busca Is Nothing Then
    
        Primeira_Ocorrencia = Busca.Address
        Resultados = Busca.Row  'Lista o primeiro resultado na variavel
    
        'Neste loop, pesquisa todas as próximas ocorrências para
        'o termo pesquisado
        Do
            Set Busca = Sheet3.Cells.FindNext(After:=Busca)
        
            'Condicional para não listar o primeiro resultado
            'pois já foi listado acima
            If Not Busca.Address Like Primeira_Ocorrencia Then
                Resultados = Resultados & ";" & Busca.Row
            End If
        Loop Until Busca.Address Like Primeira_Ocorrencia
    
    
        MatrizResultados = Split(Resultados, ";")
        
        'Atualiza dados iniciais no formulário
            SpinButton1.Max = UBound(MatrizResultados)  'Valor maximo do seletor de registros
        
        'habilita o seletor de registro
                    SpinButton1.Enabled = True
        
        'indicador do seletor de registros
                        Label_Registros_Contador.Caption = "1 de " & UBound(MatrizResultados) + 1
        
        
        'Box com o conteudo encontrado
        caixa_ppid.Text = Sheet3.Cells(MatrizResultados(0), 2).Value
        caixa_modelo.Text = Sheet3.Cells(MatrizResultados(0), 3).Value
        'caixa_link.Text = Sheet3.Cells(MatrizResultados(0), 3).Value
                 
                 
    Else    'Caso nada tenha sido encontrado, exibe mensagem informativa
    
        SpinButton1.Enabled = False     'desabilita o seletor de registros
        
            Label_Registros_Contador.Caption = ""   'zera os resultados encontrados
        
                'limpa os campos do formulário
                caixa_ppid.Text = ""
                caixa_modelo.Text = ""
        
        MsgBox "Nenhum resultado para '" & TermoPesquisado & "' foi encontrado."
        
    End If
    
Call noOpt

End Sub

Private Sub caixa_ppid_Change()

End Sub

Private Sub TextBox1_Change()

End Sub

Private Sub txt_Procurar_Change()

End Sub

Private Sub UserForm_Initialize()

Call Apresentar_on


Call Opt

' preenchendo as Combo Box - Dinamicamente

lin = 2
    Do Until Sheets("Aux_1").Cells(lin, 1) = ""
    combo_modelo.AddItem Sheets("Aux_1").Cells(lin, 1)
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


'cancela o botão de fechamento do Userform
Private Sub UserForm_QueryClose _
  (Cancel As Integer, CloseMode As Integer)
  
Call Opt

    If CloseMode = vbFormControlMenu Then
        Cancel = True
    End If
    
Call noOpt

End Sub
