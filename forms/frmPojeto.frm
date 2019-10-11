VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmPojeto 
   Caption         =   "Projeto"
   ClientHeight    =   6750
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   5190
   OleObjectBlob   =   "frmPojeto.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmPojeto"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdCadastrar_Click()
On Error GoTo cmdCadastrar_err
Dim strMSG As String: strMSG = "Favor Preencher campo "
Dim strTitulo As String: strTitulo = "CAMPO OBRIGATORIO!"

Dim strBanco As String: strBanco = Range(BancoLocal)
Dim strSheet As String: strSheet = ActiveSheet.Name
Dim strGerente As String: strGerente = Range(GerenteDeContas)
Dim Cadastro As Boolean: Cadastro = False

''' LINHA
If Me.cboLinha <> "" Then
    
    Cadastro = True
Else
    MsgBox strMSG, vbCritical + vbOKOnly, strTitulo
    Me.cboLinha.SetFocus
    Exit Sub
End If

''' FASCICULOS
If Me.txtFasciculos <> "" Then
    
    Cadastro = True
Else
    MsgBox strMSG, vbCritical + vbOKOnly, strTitulo
    Me.txtFasciculos.SetFocus
    Exit Sub
End If

''' VENDAS
If Me.cboVendas <> "" Then
    
    Cadastro = True
Else
    MsgBox strMSG, vbCritical + vbOKOnly, strTitulo
    Me.cboVendas.SetFocus
    Exit Sub
End If

''' IDIOMAS
If Me.cboIdiomas <> "" Then
    
    Cadastro = True
Else
    MsgBox strMSG, vbCritical + vbOKOnly, strTitulo
    Me.cboIdiomas.SetFocus
    Exit Sub
End If

''' TIRAGEM
If Me.txtTiragem <> "" Then
    
    Cadastro = True
Else
    MsgBox strMSG, vbCritical + vbOKOnly, strTitulo
    Me.txtTiragem.SetFocus
    Exit Sub
End If

''' ESPECIFICAÇÃO
If Me.txtEspecificacao <> "" Then
    
    Cadastro = True
Else
    MsgBox strMSG, vbCritical + vbOKOnly, strTitulo
    Me.txtEspecificacao.SetFocus
    Exit Sub
End If

''' MOEDA
If Me.cboMoeda <> "" Then
    
    Cadastro = True
Else
    MsgBox strMSG, vbCritical + vbOKOnly, strTitulo
    Me.cboMoeda.SetFocus
    Exit Sub
End If

''' ROYALTY (%)
If Me.txtRoyalty_Percentual <> "" Then
    
    Cadastro = True
Else
    MsgBox strMSG, vbCritical + vbOKOnly, strTitulo
    Me.txtRoyalty_Percentual.SetFocus
    Exit Sub
End If

''' ROYALTY (VALOR)
If Me.txtRoyalty_Valor <> "" Then
    
    Cadastro = True
Else
    MsgBox strMSG, vbCritical + vbOKOnly, strTitulo
    Me.txtRoyalty_Valor.SetFocus
    Exit Sub
End If

''' RE-IMPRESSÃO
If Me.txtReImpressao <> "" Then
    
    Cadastro = True
Else
    MsgBox strMSG, vbCritical + vbOKOnly, strTitulo
    Me.txtReImpressao.SetFocus
    Exit Sub
End If

If Cadastro Then
        
    If ProjetoAtual = "G" Or ProjetoAtual = "H" Or ProjetoAtual = "I" Or ProjetoAtual = "J" Then
    
    Else
        Range(ProjetoAtual & "13") = Me.cboLinha
        Range(ProjetoAtual & "14") = Me.txtFasciculos
    End If
    
    Range(ProjetoAtual & "15") = Me.cboVendas
    Range(ProjetoAtual & "17") = Me.cboIdiomas
    Range(ProjetoAtual & "18") = Me.txtTiragem
    Range(ProjetoAtual & "19") = Me.txtEspecificacao
    Range(ProjetoAtual & "20") = Me.cboMoeda
    Range(ProjetoAtual & "21") = Me.txtRoyalty_Percentual
    Range(ProjetoAtual & "22") = Me.txtRoyalty_Valor
    Range(ProjetoAtual & "23") = Me.txtReImpressao
        
End If


cmdCadastrar_Fim:
        
    Call cmdFechar_Click
        
    Exit Sub
cmdCadastrar_err:
    MsgBox Err.Description, vbCritical + vbOKOnly, "Cadastro de projeto"
    Resume cmdCadastrar_Fim


End Sub

Private Sub cmdFechar_Click()
    Unload Me
End Sub


Private Sub UserForm_Initialize()
Dim cPart As Range
Dim cLoc As Range
Dim strUsuario As String: strUsuario = Range(NomeUsuario)

Dim wsApoio As Worksheet
Set wsApoio = Worksheets("Apoio")

Dim wsPrincipal As Worksheet
Set wsPrincipal = Worksheets(ActiveSheet.Name)

    
    If ProjetoAtual = "" Then
    
        ProjetoAtual = "C"
        
    Else
    
        ' LINHA
        For Each cLoc In wsPrincipal.Range("LINHA")
          With Me.cboLinha
            .AddItem cLoc.value
          End With
        Next cLoc
        
        Me.cboLinha = Range(ProjetoAtual & "13")
        
        Me.txtFasciculos = Range(ProjetoAtual & "14")
        
        ''' CARREGAR COMBO BOX DE VENDAS
        ComboBoxUpdate wsPrincipal.Name, "VENDAS", Me.cboVendas
        Me.cboVendas = Range(ProjetoAtual & "15")
        
        ''' CARREGAR COMBO BOX DE IDIOMAS
        ComboBoxUpdate wsApoio.Name, "IDIOMAS", Me.cboIdiomas
        
        Me.cboIdiomas = Range(ProjetoAtual & "17")
        Me.txtTiragem = Range(ProjetoAtual & "18")
        Me.txtEspecificacao = Range(ProjetoAtual & "19")
        Me.cboVendas.SetFocus
        
        ''' CARREGAR COMBO BOX DE MOEDA
        ComboBoxUpdate wsPrincipal.Name, "MOEDA", Me.cboMoeda
        Me.cboMoeda = Range(ProjetoAtual & "20")
        
        Me.txtRoyalty_Percentual = Range(ProjetoAtual & "21")
        Me.txtRoyalty_Valor = Range(ProjetoAtual & "22")
        Me.txtReImpressao = Range(ProjetoAtual & "23")

    End If
    
End Sub

