VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserLogin 
   Caption         =   "Placeholder"
   ClientHeight    =   5640
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   5250
   OleObjectBlob   =   "UserLogin.frx":0000
   StartUpPosition =   2  'CenterScreen
End
Attribute VB_Name = "UserLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub UserForm_Initialize() ' rotina para levar o cursor ao campo Usuário

    ActiveWindow.WindowState = xlNormal
        btn_consulta_dados.SetFocus

End Sub
Private Sub botao_login_Click() 'criação de logins com perfis diferenciados

    Call Opt
        If caixa_usuario.Text = "admin" And caixa_senha.Text = "710243" Then
            MsgBox "Acesso Liberado", vbExclamation, "Seja bem vindo ao nosso sistema!"
                Unload UserLogin
                    Call Apresentar_off
                        Application.Visible = True
                        Plan2.Visible = xlSheetVisible
                        Sheet2.Visible = xlSheetVeryHidden
            
            ElseIf caixa_usuario.Text = "placeholder" And caixa_senha.Text = "placeholder" Then
                MsgBox "Acesso Liberado", vbExclamation, "Seja bem vindo, placeholder!"
                Unload UserLogin
                    Call Apresentar_off
                        Application.Visible = True
                        Plan2.Visible = xlSheetVisible
                        Sheet2.Visible = xlSheetVeryHidden
                        
          Else
              MsgBox "A senha ou o usuário estão incorretos, tente novamente", vbExclamation, " ERRO "
                caixa_usuario.SetFocus
                caixa_usuario.Text = ""
                caixa_senha.Text = ""
        End If
            Call noOpt
            
End Sub
Private Sub btn_consulta_dados_Click() 'botao para abrir a tela de double check
    
    Call copyPaste
    Unload UserLogin
        Application.Visible = False
            Call Apresentar_on
                ConsultaChecksum.Show
    Unload UserLogin
    
End Sub
Private Sub caixa_usuario_Change() 'caixa para se inserir o user para acesso ao modo DEV

End Sub
Private Sub caixa_senha_Change() 'caixa para se inserir a senha para acesso ao modo DEV

End Sub
Private Sub botao_cancelar_Click() 'botão sair, para fechar o sistema

    Call Apresentar_off
        Application.Visible = True
        Application.DisplayAlerts = False
            Worksheets("Menu").Visible = 1
               Plan1.Visible = xlSheetVeryHidden
               Plan2.Visible = xlSheetVeryHidden
        Application.DisplayFormulaBar = True
    ActiveWorkbook.Close
    
End Sub
' Desabilita o botão "x" - Fechar - do Userform
Private Sub UserForm_QueryClose _
  (Cancel As Integer, CloseMode As Integer)
        If CloseMode = vbFormControlMenu Then
            Cancel = True
        End If
    
End Sub
Private Sub usuario_Click()

End Sub
Private Sub Image2_Click()

End Sub

Private Sub Label1_Click()

End Sub

Private Sub UserForm_Click()

End Sub
