VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmEnviar 
   Caption         =   "Envio de Orçamento(s)"
   ClientHeight    =   7725
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   11025
   OleObjectBlob   =   "frmEnviar.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmEnviar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public strPesquisar As String
Public strSQL As String
Public strUsuarios As String




Private Sub cmdEnviar_Click()
On Error GoTo cmdEnviar_err

Dim strNomeUsuario As String: strNomeUsuario = Range(NomeUsuario)
Dim strBancoOrigem As String: strBancoOrigem = Range(BancoLocal)
Dim strBancoDestino As String: strBancoDestino = pathWorkSheetAddress & Controle & "_db" & "TRANSITO" & ".mdb"
Dim intCurrentRow As Integer
Dim Matriz As Variant
Dim strAssunto As String
Dim strArquivo As String
Dim strConteudo As String
Dim strSelecao As String

Dim strMSG As String
Dim strTitulo As String

Matriz = Array()

    If ListBoxChecarSelecao(Me, Me.lstOrigem.Name) = False Then
        strMSG = "Ops!!! " & Chr(10) & Chr(13) & Chr(13)
        strMSG = strMSG & "Você esqueceu de selecionar um ORÇAMENTO da lista. " & Chr(10) & Chr(13) & Chr(13)
        strTitulo = "SELEÇÃO DE ORÇAMENTO(S)!"
        
        MsgBox strMSG, vbInformation + vbOKOnly, strTitulo
    ElseIf ListBoxChecarSelecao(Me, Me.lstEmails.Name) = False Then
    
        strMSG = "Ops!!! " & Chr(10) & Chr(13) & Chr(13)
        strMSG = strMSG & "Você esqueceu de selecionar um E-MAIL da lista. " & Chr(10) & Chr(13) & Chr(13)
        strTitulo = "SELEÇÃO DE E-MAIL!"
        
        MsgBox strMSG, vbInformation + vbOKOnly, strTitulo
    Else
    
        ''' BANCO DE TRANSITO
        ''' SE O BANCO DE DESTINO JÁ EXISTIR DELETA
        If Dir(strBancoDestino) <> "" Then Kill strBancoDestino
        
        ''' CRIA BASE DE DADOS PARA EXPORTAÇÃO DE DADOS
        CriarBancoParaExportacao strBancoDestino
        
        ''' CRIAR TABELA(S) EM BASE DE DADOS DE EXPORTAÇÃO
        CriarTabelaEmBancoParaExportacao strBancoOrigem, strBancoDestino, "Orcamentos"
        CriarTabelaEmBancoParaExportacao strBancoOrigem, strBancoDestino, "OrcamentosAnexos"
        CriarTabelaEmBancoParaExportacao strBancoOrigem, strBancoDestino, "OrcamentosCustos"
        
        ''' EXPORTA PARA O BANCO ITENS SELECIONADOS
        For intCurrentRow = 0 To Me.lstOrigem.ListCount - 1
            DoEvents
        
            If Me.lstOrigem.Selected(intCurrentRow) Then
                ''' CARREGA VARIAVEL DE CONTEUDO PARA CORPO DO E-MAIL
                strConteudo = strConteudo & Me.lstOrigem.List(intCurrentRow) & Chr(13)
                
                ''' CARREGA MATRIZ
                Matriz = Split(Me.lstOrigem.List(intCurrentRow), " - ")
                
                ''' EXPORTA ITEM SELECIONADO PARA BANCO
                ExportarOrcamento strBancoOrigem, strBancoDestino, CStr(Matriz(0)), CStr(Matriz(2))
                
                ''' CARREGA VARIAVEL DE SELEÇÃO DE USUÁRIOS
                strSelecao = strSelecao & "'" & CStr(Matriz(2)) & "',"
                
                ''' DESMARCAR ITEM SELECIONADO
                Me.lstOrigem.Selected(intCurrentRow) = False
            End If
        
        Next intCurrentRow
            
        strSelecao = Left(strSelecao, Len(strSelecao) - 1) & ""
        
        ''' COMPACTA BASE DE DADOS
'        Zip strBancoDestino, Left(strBancoDestino, Len(strBancoDestino) - 3) & "zip"
        
        
        Compact strBancoDestino, Left(strBancoDestino, Len(strBancoDestino) - 3) & "zip"
        
        ''' DELETA BASE DE DADOS TEMPORARIA
        If Dir$(strBancoDestino) <> "" Then Kill strBancoDestino
        
        ''' CARREGA VARIAVEIS RESPONSAVEIS PELO ENVIO DO E-MAIL
        strAssunto = Controle & "_" & "TRANSITO"
        strArquivo = Left(strBancoDestino, Len(strBancoDestino) - 3) & "zip"
    
    
        ''' EXPORTA PARA O BANCO ITENS SELECIONADOS
        For intCurrentRow = 0 To Me.lstEmails.ListCount - 1
            DoEvents
        
            If Me.lstEmails.Selected(intCurrentRow) Then
               
                ''' CARREGA MATRIZ
                Matriz = Split(Me.lstEmails.List(intCurrentRow), " - ")
                
                ''' ENVIO
                EnviarOrcamentos CStr(Matriz(0)), strAssunto, strArquivo, strConteudo
                
                ''' DESMARCAR ITEM SELECIONADO
                Me.lstEmails.Selected(intCurrentRow) = False
            End If
        
        Next intCurrentRow
        
        ''' DELETA BASE DE DADOS COMPACTADO
        If Dir$(strArquivo) <> "" Then Kill strArquivo
        
        MsgBox "Envio concluído", vbInformation + vbOKOnly, "Envio de Orçamento(s)"
        

    End If

    


cmdEnviar_Fim:
        
    
    
    Exit Sub
cmdEnviar_err:
    MsgBox Err.Description, vbCritical + vbOKOnly, "Envio de Orçamento(s)"
    Resume cmdEnviar_Fim

End Sub

Private Sub EnviarPorDePartamento(ByVal strDepartamento As String, strAssunto As String, strArquivo As String, strConteudo As String)
'' CAMINHO DO BANCO
Dim strBancoOrigem As String: strBancoOrigem = Range(BancoLocal)

'' POSICIONA O BANCO DE ORIGEM
Dim dbOrigem As DAO.Database
Set dbOrigem = DBEngine.OpenDatabase(strBancoOrigem, False, False, "MS Access;PWD=" & SenhaBanco)

'' SELECIONA OS REGISTROS DA ORIGEM
Dim rstOrigem As DAO.Recordset
Set rstOrigem = dbOrigem.OpenRecordset("Select * from qryUsuarios where DPTO = '" & strDepartamento & "'")

While Not rstOrigem.EOF

    EnviarOrcamentos rstOrigem.Fields("eMail"), strAssunto, strArquivo, strConteudo
    rstOrigem.MoveNext

Wend

rstOrigem.Close
dbOrigem.Close


Set rstOrigem = Nothing
Set dbOrigem = Nothing

End Sub


Private Sub EnviarPorDeVendedores(ByVal strSelecao As String, strAssunto As String, strArquivo As String, strConteudo As String)
'' CAMINHO DO BANCO
Dim strBancoOrigem As String: strBancoOrigem = Range(BancoLocal)

'' POSICIONA O BANCO DE ORIGEM
Dim dbOrigem As DAO.Database
Set dbOrigem = DBEngine.OpenDatabase(strBancoOrigem, False, False, "MS Access;PWD=" & SenhaBanco)

'' SELECIONA OS REGISTROS DA ORIGEM
Dim rstOrigem As DAO.Recordset
Set rstOrigem = dbOrigem.OpenRecordset("SELECT qryUsuarios.eMail FROM qryUsuarios WHERE qryUsuarios.Usuario In (" & strSelecao & ") AND (qryUsuarios.ExclusaoVirtual)=False")

While Not rstOrigem.EOF

    EnviarOrcamentos rstOrigem.Fields("eMail"), strAssunto, strArquivo, strConteudo
    rstOrigem.MoveNext

Wend

rstOrigem.Close
dbOrigem.Close


Set rstOrigem = Nothing
Set dbOrigem = Nothing

End Sub
Private Sub cmdPesquisar_Click()
Dim strBanco As String: strBanco = Range(BancoLocal)

Dim retValor As Variant

    retValor = InputBox("Digite uma palavra para fazer o filtro:", "Filtro", strPesquisar, 0, 0)
    strPesquisar = retValor
    
    MontarPesquisa
    
    ListBoxCarregar strBanco, Me, Me.lstOrigem.Name, "Pesquisa", strSQL
    Me.Repaint

End Sub

Private Sub cmdTodos_Click()
Dim intCurrentRow As Integer
            
For intCurrentRow = 0 To Me.lstOrigem.ListCount - 1
    If Not IsNull(Me.lstOrigem.Column(0, intCurrentRow)) Then
        Me.lstOrigem.Selected(intCurrentRow) = True
    End If
Next intCurrentRow

End Sub

Private Sub cmdNenhum_Click()
Dim intCurrentRow As Integer
            
For intCurrentRow = 0 To Me.lstOrigem.ListCount - 1
    If Not IsNull(Me.lstOrigem.Column(0, intCurrentRow)) Then
        Me.lstOrigem.Selected(intCurrentRow) = False
    End If
Next intCurrentRow

End Sub


Private Sub UserForm_Initialize()
Dim strBanco As String: strBanco = Range(BancoLocal)
Dim sqlUsuarios As String: strUsuarios = Range(NomeUsuario)
Dim sqlEnvio As String

MontarPesquisa
sqlUsuarios = "Select * from qryUsuarios WHERE (((qryUsuarios.ExclusaoVirtual)=No)) Order By Usuario"

sqlEnvio = "SELECT * From qryUsuarios WHERE (((qryUsuarios.Usuario) In (Select Usuarios from qryUsuariosUsuarios where Usuario = '" & strUsuarios & "')) AND ((qryUsuarios.ExclusaoVirtual)=No)) ORDER BY qryUsuarios.Usuario Union SELECT * FROM qryUsuarios WHERE (((qryUsuarios.DPTO) In ('Produção','FINANCEIRO')))"

ListBoxCarregar strBanco, Me, Me.lstOrigem.Name, "Pesquisa", strSQL
ListBoxCarregar strBanco, Me, Me.lstEmails.Name, "email", sqlEnvio


End Sub


Private Sub MontarPesquisa()

strSQL = "SELECT qryOrcamentosEnviar.Pesquisa FROM qryOrcamentosEnviar WHERE ((qryOrcamentosEnviar.Pesquisa) Like '*" & strPesquisar & "*')"
strSQL = strSQL + " AND ((qryOrcamentosEnviar.VENDEDOR) In (Select Descricao01 from admCategorias where codRelacao = (SELECT admCategorias.codCategoria FROM admCategorias WHERE ((admCategorias.Categoria)='" & strUsuarios & "')) and Categoria = 'Usuarios'))"
strSQL = strSQL + " AND ((qryOrcamentosEnviar.DEPARTAMENTO) In (Select Descricao01 from admCategorias where codRelacao = (SELECT admCategorias.codCategoria FROM admCategorias WHERE ((admCategorias.Categoria)='" & strUsuarios & "')) and Categoria = 'Departamentos')) "
strSQL = strSQL + "ORDER BY qryOrcamentosEnviar.CONTROLE DESC , qryOrcamentosEnviar.VENDEDOR"


End Sub
