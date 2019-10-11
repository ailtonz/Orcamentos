VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmEtapas 
   Caption         =   "Seleção de Proxima Etapa"
   ClientHeight    =   2760
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5130
   OleObjectBlob   =   "frmEtapas.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmEtapas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdStatus_Click()
Dim BaseDeDados As String: BaseDeDados = Range(BancoLocal)
Dim strControle As String: strControle = ActiveSheet.Name
Dim strVendedor As String: strVendedor = Range(GerenteDeContas)

Dim strMSG As String
Dim strTitulo As String

    If IsNull(Me.lstEtapas.value) Then
        strMSG = "ATENÇÃO: Por favor selecione um item da lista. " & Chr(10) & Chr(13) & Chr(13)
        strTitulo = "Proxima Etapa"
        
        MsgBox strMSG, vbInformation + vbOKOnly, strTitulo
    Else
        admOrcamentoAtualizarEtapa BaseDeDados, strControle, strVendedor, CodigoEtapa(BaseDeDados, Me.lstEtapas.value)
        Me.Hide
    End If
    
End Sub

Private Sub lstEtapas_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    cmdStatus_Click
End Sub

Private Sub UserForm_Activate()
Dim BaseDeDados As String: BaseDeDados = Range(BancoLocal)
Dim strUsuario As String: strUsuario = Range(NomeUsuario)

'ListBoxCarregar BaseDeDados, Me, Me.lstEtapas.Name, "Status", "Select Status from qryEtapas order by Atual"

ListBoxCarregar BaseDeDados, Me, Me.lstEtapas.Name, "Etapa", "Select Etapa from qryUsuariosVoltarEtapa where usuario = '" & strUsuario & "' order by Atual"

'qryUsuariosVoltarEtapa


'Dim strBanco(2) As String
'Dim strStatus
'
'Me.lstEtapas.Clear
'
'strBanco(0) = ("CANCELADO")
'strBanco(1) = ("DESCONTO")
'strBanco(2) = ("RE-CUSTO")
'
'strStatus = Array()
'
'For Each strStatus In strBanco
'
'    Me.lstEtapas.AddItem strStatus
'
'Next

End Sub

