VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ConsultaChecksum 
   Caption         =   "Placeholder"
   ClientHeight    =   5970
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   12075
   OleObjectBlob   =   "ConsultaChecksum.frx":0000
End
Attribute VB_Name = "ConsultaChecksum"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' Declaração de variáveis publicas universais
Public MatrizResultados As Variant
Public Total_Ocorrencias As Long

Private Sub Image4_BeforeDragOver(ByVal Cancel As MSForms.ReturnBoolean, ByVal Data As MSForms.DataObject, ByVal X As Single, ByVal Y As Single, ByVal DragState As MSForms.fmDragState, ByVal Effect As MSForms.ReturnEffect, ByVal Shift As Integer)

End Sub

Private Sub Label3_Click()

End Sub

Private Sub UserForm_Initialize() ' inicialização da tela de consulta - Double Check

    Call Apresentar_on
    Call Opt
    ' preenchendo as Combo Box - Dinamicamente
        lin = 2
            Do Until Sheets("Aux_1").Cells(lin, 1) = ""
                combo_modelo.AddItem Sheets("Aux_1").Cells(lin, 1)
        lin = lin + 1
            Loop
    
    'rotinas para navegar entre os resultados
        SpinButton1.Enabled = False
            Label_Registros_Contador.Caption = ""
                Call noOpt
    
End Sub
Private Sub botao_pesquisaModelo_Click() ' rotina para inicializar o botão DOUBLE CHECK

Call Opt
    If Me.combo_modelo.Text = "" Then
        MsgBox "ESCANEIE O CHECKSUM CORRESPONDENTE AO QUE FOI GRAVADO", vbInformation, "ALERTA"
        
    'limpa os campos do formulário
    caixa_ppid.Text = ""
    caixa_modelo.Text = ""
    caixa_modelo_mb.Text = ""
    caixa_rev_grav.Text = ""
           
    Else
        Call ProcuraPersonalizada(Me.combo_modelo.Text)
    End If
Call noOpt

End Sub
Private Sub Image2_Click() ' rotina para sair da tela de consulta com segurança e voltar a tela de login

    Application.DisplayAlerts = False
        Sheet2.Visible = xlSheetVisible
        Plan2.Visible = xlSheetVeryHidden
        Sheet3.Visible = xlSheetVeryHidden
            Unload Me
            UserLogin.Show

End Sub
Private Sub ProcuraPersonalizada(ByVal TermoPesquisado As String) 'Executa a busca do CHECKSUM no Banco de Dados

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
    
        'Neste loop, pesquisa todas as próximas ocorrências para o termo pesquisado
        Do
            Set Busca = Sheet3.Cells.FindNext(After:=Busca)
        
        'Condicional para não listar o primeiro resultado pois já foi listado acima
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
            'caixa_ppid.Text = Sheet3.Cells(MatrizResultados(0), 1).Value 'Box com o CHECKSUM encontrado
            caixa_ppid.Text = "CHECKSUM PASS - SEGUIR PROCESSO"
            caixa_modelo.Text = Sheet3.Cells(MatrizResultados(0), 2).Value 'Box com o PN COMPONENTE GRAVADO encontrado
            caixa_modelo_mb.Text = Sheet3.Cells(MatrizResultados(0), 3).Value 'Box com o MODELO MB encontrado
            caixa_rev_grav.Text = Sheet3.Cells(MatrizResultados(0), 4).Value 'Box com a REVISAO encontrada
            
        'limpa os campos do formulário
        'combo_modelo.Text = ""
        Me.combo_modelo.SetFocus
        
         'rotina para salvar registro  - condição PASS
            LogInformation combo_modelo.Text & ";" & caixa_modelo.Text & ";" & _
            caixa_modelo_mb.Text & ";" & caixa_rev_grav.Text & ";" & "PASS" & ";" & _
            Application.UserName & ";" & Format(Now, "yyyy-mm-dd hh:mm:ss")
  
    Else    'Caso nada tenha sido encontrado, exibe mensagem informativa
        SpinButton1.Enabled = False     'desabilita o seletor de registros
            Label_Registros_Contador.Caption = ""   'zera os resultados encontrados
            
                'limpa os campos do formulário
                caixa_ppid.Text = ""
                caixa_modelo.Text = ""
                caixa_modelo_mb.Text = ""
                caixa_rev_grav.Text = ""
                                
        'Mensagem que retorna ao usuário o fato do CHECKSUM estar incorreto
        MsgBox "CHECKSUM '****" & _
        TermoPesquisado & "****' INCORRETO!! ESCANEAR NOVAMENTE! SE O ERRO PERSISTIR, CHAMAR ENG. DE TESTES OU PRODUTO!!", _
        vbCritical, "ERRO - CHECKSUM"
                
        'rotina para salvar registro de trabalho  - condição FAIL
        LogInformation combo_modelo.Text & ";" & ";" & ";" & ";" & "FAIL" & ";" & _
            Application.UserName & ";" & Format(Now, "yyyy-mm-dd hh:mm:ss")
            
    'limpa os campos do formulário
    combo_modelo.Text = ""
    
    End If
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
Private Sub caixa_modelo_Change()

End Sub

Private Sub caixa_modelo_mb_Change()

End Sub
Private Sub caixa_rev_grav_Change()

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
Private Sub caixa_resultadoPPID_Change()

End Sub
Private Sub caixa_resultadoModelo_Change()

End Sub
Private Sub caixa_resultadoSemana_Change()

End Sub
Private Sub caixa_pesquisa_Change()

End Sub
Private Sub caixa_ppid_Change()

End Sub
Private Sub Label15_Click()

End Sub
Private Sub Label18_Click()

End Sub
Private Sub Label19_Click()

End Sub
Private Sub Label9_Click()

End Sub
Private Sub TextBox1_Change()

End Sub
Private Sub txt_Procurar_Change()

End Sub
