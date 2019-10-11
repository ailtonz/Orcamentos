VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "EstaPasta_de_trabalho"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True


Private Sub Workbook_BeforeClose(Cancel As Boolean)

'If Not Cancel = True Then
''    ENVIAR_DADOS
'    admOrcamentoFormulariosLimpar
'End If

GUIAS_APOIO_DESATIVAR

End Sub

Private Sub Workbook_Open()

GUIAS_APOIO_DESATIVAR


admAtualizarLocalizacaoDaBaseDeDados
admAtualizarUsuario
UpdateSystem UCase(Range(NomeUsuario))
admAtualizarGuiaDeApoio

'RECEBER_DADOS
        
End Sub

Private Sub Workbook_BeforeSave(ByVal SaveAsUI As Boolean, Cancel As Boolean)
On Error GoTo Workbook_BeforeSave_err

Dim strMSG As String
Dim strTitulo As String
Dim RetVal As Variant

Dim dbOrcamento As DAO.Database
Dim rstOrcamento As DAO.Recordset
'Dim rstIntervalos As DAO.Recordset


Dim strOrcamento As String
Dim strIntervalos As String

Dim intProxEtapa As String


Dim strControle As String: strControle = ActiveSheet.Name
Dim strVendedor As String: strVendedor = Range(GerenteDeContas)

Dim BaseDeDados As String: BaseDeDados = Range(BancoLocal)

'admVerificarBaseDeDados

If strControle <> "" And strVendedor <> "" And strVendedor <> strControle Then

    '   BLOQUEIO TELA
    Application.ScreenUpdating = False

    '   CARREGA BASE DE DADOS
    Set dbOrcamento = DBEngine.OpenDatabase(BaseDeDados, False, False, "MS Access;PWD=" & SenhaBanco)

    '   SELECIONA O ORÇAMENTO
    strOrcamento = "SELECT Orcamentos.* " & _
         " FROM Orcamentos  " & _
         " WHERE (((CONTROLE)='" & strControle & "') AND ((VENDEDOR)= '" & strVendedor & "')) "
    Set rstOrcamento = dbOrcamento.OpenRecordset(strOrcamento)

'    '   LISTA OS INTERVALOS
'    strIntervalos = "Select * from qryEtapasIntervalosEdicoes where Departamento = '" & _
'                    rstOrcamento.Fields("Departamento") & "' and Status = '" & rstOrcamento.Fields("Status") & "'"
'    Set rstIntervalos = dbOrcamento.OpenRecordset(strIntervalos)

    '   ENVIAR ORÇAMENTO PARA PROXIMA ETAPA
    intProxEtapa = admProxEtapa(BaseDeDados, rstOrcamento.Fields("Departamento"), rstOrcamento.Fields("Status"))

    '   SE A PROXIMA ETAPA FOR DIFERENTE DE 1(UM) MANTEM OS PROCESSOS. AGORA -->
    '   SE A PROXIMA ETAPA FOR IGUAL A 1(UM) ENTÃO DESCONSIDERA TODAS AS OPERAÇÕES.
    If intProxEtapa <> 1 Then

        '   QUESTIONA TERMINIO DE ETAPA
        strMSG = "ATENÇÃO: Esta etapa está finalizada?. " & Chr(10) & Chr(13) & Chr(13) & _
                 "Click em SIM caso deseje enviar o Orçamento para a proxima etapa." & Chr(10) & Chr(13) & Chr(13) & _
                 "Click em NÃO caso deseje continuar alterando no futuro."
        strTitulo = "Fluxo de etapas dos orçamentos"
        RetVal = MsgBox(strMSG, vbQuestion + vbYesNo, strTitulo)

        If RetVal = vbYes Then

            If intProxEtapa = 0 Then

                If Range("K12").value = "OK" Then

                    strMSG = "Deseja efetuar a venda?"
                    RetVal = MsgBox(strMSG, vbQuestion + vbYesNo, strTitulo)
                    If RetVal = vbYes Then
                        admOrcamentoAtualizarEtapa BaseDeDados, strControle, strVendedor, intProxEtapa
                    ElseIf RetVal = vbNo Then
                        frmEtapas.Show
                    End If

                Else
                    strMSG = "ATENÇÃO: VENDA NÃO AUTORIZADA!!! "
                    strTitulo = "VENDA NÃO AUTORIZADA"

                    MsgBox strMSG, vbQuestion + vbOKOnly, strTitulo

                    'GoTo Workbook_BeforeSave_Fim

                    Exit Sub

                End If


            Else
                admOrcamentoAtualizarEtapa BaseDeDados, strControle, strVendedor, intProxEtapa
            End If

            admNotificacoes BaseDeDados, rstOrcamento.Fields("CONTROLE"), rstOrcamento.Fields("VENDEDOR"), rstOrcamento.Fields("Status")

        End If

'        ' SALVAR CADASTROS DISPONIVEIS
'        While Not rstIntervalos.EOF
                    
            ''CUSTOS
            If UsuarioAdministrativo(BaseDeDados, Range(NomeUsuario)) Then
                CadastroOrcamentoCustos BaseDeDados, strControle, strVendedor
            Else
                CadastroOrcamento BaseDeDados, strControle, strVendedor
                CadastroOrcamentoImpressao BaseDeDados, strControle, strVendedor
            End If
        
        

'            '' ORÇAMENTO
'            If rstIntervalos.Fields("Intervalo") = "ORÇAMENTO" Then
'                CadastroOrcamento BaseDeDados, strControle, strVendedor
'            '' IMPRESSÃO
'            ElseIf rstIntervalos.Fields("Intervalo") = "IMPRESSÃO" Then
'                CadastroOrcamentoImpressao BaseDeDados, strControle, strVendedor
'            ''CUSTOS
'            ElseIf rstIntervalos.Fields("Intervalo") = "CUSTOS" Then
'                CadastroOrcamento BaseDeDados, strControle, strVendedor
'                CadastroOrcamentoCustos BaseDeDados, strControle, strVendedor
'            End If

'            rstIntervalos.MoveNext
'
'        Wend


    End If

    rstOrcamento.Close
'    rstIntervalos.Close
    dbOrcamento.Close

    '   LIMPAR FORMULARIOS
    DesbloqueioDeGuia SenhaBloqueio
    admOrcamentoFormulariosLimpar

    '   RENOMEAR GUIA DE CONTROLE
    ActiveSheet.Name = Range(NomeUsuario)

    BloqueioDeGuia SenhaBloqueio

    '   POSICIONA CURSOR
    Range(InicioCursor).Select

    '   LIBERAR TELA
    Application.ScreenUpdating = True

End If

Workbook_BeforeSave_Fim:
    Set dbOrcamento = Nothing
    Set rstOrcamento = Nothing
'    Set rstIntervalos = Nothing
    Set rstNotificacoes = Nothing

    Exit Sub
Workbook_BeforeSave_err:

    MsgBox Err.Description
    Resume Workbook_BeforeSave_Fim


End Sub

Sub testeDPTO()
Dim BaseDeDados As String: BaseDeDados = Range(BancoLocal)

MsgBox UsuarioAdministrativo(BaseDeDados, "fabiana")

End Sub
