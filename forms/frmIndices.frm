VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmIndices 
   Caption         =   "Indices de cálculos"
   ClientHeight    =   6000
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   6705
   OleObjectBlob   =   "frmIndices.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmIndices"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdAtualizarLinha_Click()
Dim strMSG As String
Dim strTitulo As String

If ListBoxChecarSelecao(Me, Me.lstLinha.Name) = False Then
    strMSG = "Ops!!! " & Chr(10) & Chr(13) & Chr(13)
    strMSG = strMSG & "Você esqueceu de selecionar um item da lista. " & Chr(10) & Chr(13) & Chr(13)
    strTitulo = "Atualização da linha de produto!"
    
    MsgBox strMSG, vbInformation + vbOKOnly, strTitulo
Else
    
    Dim strBanco As String: strBanco = Range(BancoLocal)
    Dim strUsuario As String: strUsuario = Range(GerenteDeContas)
    Dim strControle As String: strControle = ActiveSheet.Name
    Dim strPropriedade As String: strPropriedade = "LINHA"
    Dim strIndice As String: strIndice = Trim(DivisorDeTexto(Me.lstLinha.value, "|", 0))
    Dim strValor_01 As String: strValor_01 = IIf((Me.txtLinhaValor01) = "", 0, Me.txtLinhaValor01)
    Dim strValor_02 As String: strValor_02 = IIf((Me.txtLinhaValor02) = "", 0, Me.txtLinhaValor02)
    Dim strSQL As String
    
    
    ''' GERENCIAR INDICE DE LINHA DE PRODUTOS
    admGerenciarIndicesDeCalculos _
                strBanco, _
                strUsuario, _
                strControle, _
                strPropriedade, _
                strIndice, _
                strValor_01, _
                strValor_02
    
    ''' ATUALIZAR LISTAGEM
    Me.lstLinha.Clear
    strSQL = _
    "Select * from qryOrcamentosIndicesDeCalculos where Vendedor = '" & strUsuario & "' AND Controle = '" & strControle & "'  AND Propriedade =  '" & strPropriedade & "'"
    
    ListBoxCarregar strBanco, Me, Me.lstLinha.Name, "strDescricao", strSQL
        
    ''' LIMPAR LINHA DE PRODUTOS
    Application.ScreenUpdating = False
    DesbloqueioDeGuia SenhaBloqueio
    
    Range("L3:N3").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.ClearContents
        
    CarregarAnexoLinha strBanco, strControle, strUsuario, 3, 12
               
    BloqueioDeGuia SenhaBloqueio
    Application.ScreenUpdating = True
    Range(InicioCursor).Select
    
                
End If
                
End Sub

Private Sub cmdAtualizarMoeda_Click()
Dim strMSG As String
Dim strTitulo As String

If ListBoxChecarSelecao(Me, Me.lstMoeda.Name) = False Then
    strMSG = "Ops!!! " & Chr(10) & Chr(13) & Chr(13)
    strMSG = strMSG & "Você esqueceu de selecionar um item da lista. " & Chr(10) & Chr(13) & Chr(13)
    strTitulo = "Atualização da moeda!"
    
    MsgBox strMSG, vbInformation + vbOKOnly, strTitulo
Else
    
    Dim strBanco As String: strBanco = Range(BancoLocal)
    Dim strUsuario As String: strUsuario = Range(GerenteDeContas)
    Dim strControle As String: strControle = ActiveSheet.Name
    Dim strPropriedade As String: strPropriedade = "MOEDA"
    Dim strIndice As String: strIndice = Trim(DivisorDeTexto(Me.lstMoeda.value, "|", 0))
    Dim strValor_01 As String: strValor_01 = IIf((Me.txtMoedaValor01) = "", 0, Me.txtMoedaValor01)
    Dim strValor_02 As String: strValor_02 = ""
    Dim strSQL As String
    
    
    ''' GERENCIAR INDICE DE LINHA DE PRODUTOS
    admGerenciarIndicesDeCalculos _
                strBanco, _
                strUsuario, _
                strControle, _
                strPropriedade, _
                strIndice, _
                strValor_01, _
                strValor_02
    
    ''' ATUALIZAR LISTAGEM
    Me.lstMoeda.Clear
    strSQL = _
    "Select * from qryOrcamentosIndicesDeCalculos where Vendedor = '" & strUsuario & "' AND Controle = '" & strControle & "'  AND Propriedade =  '" & strPropriedade & "'"
    
    ListBoxCarregar strBanco, Me, Me.lstMoeda.Name, "strDescricao", strSQL
    
    
    ''' LIMPAR MOEDA
    Application.ScreenUpdating = False
    DesbloqueioDeGuia SenhaBloqueio
    
    Range("P3:Q3").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.ClearContents
        
    CarregarAnexoMoeda strBanco, strControle, strUsuario, 3, 16
               
    BloqueioDeGuia SenhaBloqueio
    Application.ScreenUpdating = True
    Range(InicioCursor).Select

                
End If

End Sub

Private Sub cmdAtualizarVenda_Click()
Dim strMSG As String
Dim strTitulo As String

If ListBoxChecarSelecao(Me, Me.lstVenda.Name) = False Then
    strMSG = "Ops!!! " & Chr(10) & Chr(13) & Chr(13)
    strMSG = strMSG & "Você esqueceu de selecionar um item da lista. " & Chr(10) & Chr(13) & Chr(13)
    strTitulo = "Atualização da Venda!"
    
    MsgBox strMSG, vbInformation + vbOKOnly, strTitulo
Else
    
    Dim strBanco As String: strBanco = Range(BancoLocal)
    Dim strUsuario As String: strUsuario = Range(GerenteDeContas)
    Dim strControle As String: strControle = ActiveSheet.Name
    Dim strPropriedade As String: strPropriedade = "VENDA"
    Dim strIndice As String: strIndice = Trim(DivisorDeTexto(Me.lstVenda.value, "|", 0))
    Dim strValor_01 As String: strValor_01 = IIf((Me.txtVendaValor01) = "", 0, Me.txtVendaValor01)
    Dim strValor_02 As String: strValor_02 = ""
    Dim strSQL As String
    
    
    ''' GERENCIAR INDICE DE LINHA DE PRODUTOS
    admGerenciarIndicesDeCalculos _
                strBanco, _
                strUsuario, _
                strControle, _
                strPropriedade, _
                strIndice, _
                strValor_01, _
                strValor_02
    
    ''' ATUALIZAR LISTAGEM
    Me.lstVenda.Clear
    
    strSQL = _
    "Select * from qryOrcamentosIndicesDeCalculos where Vendedor = '" & strUsuario & "' AND Controle = '" & strControle & "'  AND Propriedade =  '" & strPropriedade & "'"
    
    ListBoxCarregar strBanco, Me, Me.lstVenda.Name, "strDescricao", strSQL
    
    
    
    ''' LIMPAR VENDA
    Application.ScreenUpdating = False
    DesbloqueioDeGuia SenhaBloqueio
    
    Range("S3:T3").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.ClearContents
        
    CarregarAnexoVenda strBanco, strControle, strUsuario, 3, 19
               
    BloqueioDeGuia SenhaBloqueio
    Application.ScreenUpdating = True
    Range(InicioCursor).Select
                    
End If

End Sub

Private Sub cmdAtualizarDescontos_Click()
Dim strMSG As String
Dim strTitulo As String

If ListBoxChecarSelecao(Me, Me.lstDescontos.Name) = False Then
    strMSG = "Ops!!! " & Chr(10) & Chr(13) & Chr(13)
    strMSG = strMSG & "Você esqueceu de selecionar um item da lista. " & Chr(10) & Chr(13) & Chr(13)
    strTitulo = "Atualização do Desconto!"
    
    MsgBox strMSG, vbInformation + vbOKOnly, strTitulo
Else
    
    Dim strBanco As String: strBanco = Range(BancoLocal)
    Dim strUsuario As String: strUsuario = Range(GerenteDeContas)
    Dim strControle As String: strControle = ActiveSheet.Name
    Dim strPropriedade As String: strPropriedade = "DESCONTO"
    Dim strIndice As String: strIndice = Trim(DivisorDeTexto(Me.lstDescontos.value, "|", 0))
    Dim strValor_01 As String: strValor_01 = IIf((Me.txtDescontoValor01) = "", 0, Me.txtDescontoValor01)
    Dim strValor_02 As String: strValor_02 = ""
    Dim strSQL As String
    
    
    ''' GERENCIAR INDICE DE LINHA DE PRODUTOS
    admGerenciarIndicesDeCalculos _
                strBanco, _
                strUsuario, _
                strControle, _
                strPropriedade, _
                strIndice, _
                strValor_01, _
                strValor_02
    
    ''' ATUALIZAR LISTAGEM
    Me.lstDescontos.Clear
    
    strSQL = _
    "Select * from qryOrcamentosIndicesDeCalculos where Vendedor = '" & strUsuario & "' AND Controle = '" & strControle & "'  AND Propriedade =  '" & strPropriedade & "'"
    
    ListBoxCarregar strBanco, Me, Me.lstDescontos.Name, "strDescricao", strSQL
        
    
    ''' LIMPAR DESCONTOS
    Application.ScreenUpdating = False
    DesbloqueioDeGuia SenhaBloqueio
    
    Range("V3:W3").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.ClearContents
        
    CarregarAnexoDesconto strBanco, strControle, strUsuario, 3, 22
               
    BloqueioDeGuia SenhaBloqueio
    Application.ScreenUpdating = True
    Range(InicioCursor).Select
                
End If

End Sub

Private Sub lstLinha_Click()
    Me.txtLinhaValor01 = Trim(DivisorDeTexto(Me.lstLinha.value, "|", 1))
    Me.txtLinhaValor02 = Trim(DivisorDeTexto(Me.lstLinha.value, "|", 2))
End Sub

Private Sub lstMoeda_Click()
    Me.txtMoedaValor01 = Trim(DivisorDeTexto(Me.lstMoeda.value, "|", 1))
End Sub

Private Sub lstVenda_Click()
    Me.txtVendaValor01 = Trim(DivisorDeTexto(Me.lstVenda.value, "|", 1))
End Sub

Private Sub lstDescontos_Click()
    Me.txtDescontoValor01 = Trim(DivisorDeTexto(Me.lstDescontos.value, "|", 1))
End Sub

Private Sub UserForm_Initialize()
Dim strBanco As String: strBanco = Range(BancoLocal)
Dim strUsuario As String: strUsuario = Range(GerenteDeContas)
Dim strControle As String: strControle = ActiveSheet.Name
Dim strSQL As String

    strSQL = "Select * from qryOrcamentosIndicesDeCalculos where Vendedor = '" & strUsuario & "' AND Controle = '" & strControle & "'  AND Propriedade = 'LINHA'"
    ListBoxCarregar strBanco, Me, Me.lstLinha.Name, "strDescricao", strSQL
    
    strSQL = "Select * from qryOrcamentosIndicesDeCalculos where Vendedor = '" & strUsuario & "' AND Controle = '" & strControle & "'  AND Propriedade = 'MOEDA'"
    ListBoxCarregar strBanco, Me, Me.lstMoeda.Name, "strDescricao", strSQL
    
    strSQL = "Select * from qryOrcamentosIndicesDeCalculos where Vendedor = '" & strUsuario & "' AND Controle = '" & strControle & "'  AND Propriedade = 'VENDA'"
    ListBoxCarregar strBanco, Me, Me.lstVenda.Name, "strDescricao", strSQL
    
    strSQL = "Select * from qryOrcamentosIndicesDeCalculos where Vendedor = '" & strUsuario & "' AND Controle = '" & strControle & "'  AND Propriedade = 'DESCONTO'"
    ListBoxCarregar strBanco, Me, Me.lstDescontos.Name, "strDescricao", strSQL
    
    
End Sub
