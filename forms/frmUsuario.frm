VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmUsuario 
   Caption         =   "Cadastro de Usuário"
   ClientHeight    =   3990
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   7815
   OleObjectBlob   =   "frmUsuario.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmUsuario"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdSalvar_Click()
Dim strBanco As String: strBanco = Range(BancoLocal)

    If ExistenciaUsuario(Range(BancoLocal), "", Me.txtNome) Then
        admUsuarioSalvar Range(BancoLocal), "Vendas", "", Me.txtNome, Me.txtEmail, Me.txtGerenteContas, Me.txtTelefone, Me.txtCelular01, Me.txtCelular02, Me.txtIdNextel
    Else
        admUsuarioNovo Range(BancoLocal), "Vendas", "", Me.txtNome, Me.txtEmail, Me.txtGerenteContas, Me.txtTelefone, Me.txtCelular01, Me.txtCelular02, Me.txtIdNextel
    End If

    
End Sub
