Attribute VB_Name = "basDataBase"
Option Explicit

Sub CriarBancoParaExportacao(strBancoDestino As String)
Dim oAccess As Access.Application
Dim dbDestino As DAO.Database

Set oAccess = New Access.Application
Set dbDestino = DBEngine.CreateDatabase(strBancoDestino, dbLangGeneral & ";pwd=" & SenhaBanco, dbVersion40)

dbDestino.Close

Set dbDestino = Nothing
Set oAccess = Nothing

End Sub

Sub CriarTabelaEmBancoParaExportacao(strOrigem As String, strDestino As String, strTabela As String)
Dim dbOrigem As DAO.Database
Dim tbORIGEM As DAO.TableDef
Dim dbDestino As DAO.Database
Dim tdfNew As DAO.TableDef


Set dbOrigem = DBEngine.OpenDatabase(strOrigem, False, False, "MS Access;PWD=" & SenhaBanco)
Set tbORIGEM = dbOrigem.TableDefs(strTabela)
Set dbDestino = DBEngine.OpenDatabase(strDestino, False, False, "MS Access;PWD=" & SenhaBanco)
Set tdfNew = dbDestino.CreateTableDef(strTabela)

Dim x As Integer

    For x = 0 To dbOrigem.TableDefs(strTabela).Fields.Count - 1

        With tdfNew

            .Fields.Append .CreateField(dbOrigem.TableDefs(strTabela).Fields(x).Properties("name"), dbOrigem.TableDefs(strTabela).Fields(x).Properties("type"), dbOrigem.TableDefs(strTabela).Fields(x).Properties("size"))

        End With

    Next x

   dbDestino.TableDefs.Append tdfNew

'''Delete new TableDef because this is a demonstration.
'''dbDESTINO.TableDefs.Delete tdfNew.Name

   dbDestino.Close
   dbOrigem.Close

End Sub

Sub ExportarDadosTabelaOrigemParaTabelaDestino(ByVal strOrigem As String, ByVal strDestino As String, ByVal strTabela As String)
''' EXPORTAR DADOS DA TABELA ORIGEM PARA A TABELA DESTINO (AMBAS COM A MESMA EXTRUTURA)
''==============================''
''           ORIGEM
''==============================''

'' POSICIONA O BANCO DE ORIGEM
Dim dbOrigem As DAO.Database
Set dbOrigem = DBEngine.OpenDatabase(strOrigem, False, False, "MS Access;PWD=" & SenhaBanco)

'' SELECIONA A TABELA DE ORIGEM
Dim tbORIGEM As DAO.TableDef
Set tbORIGEM = dbOrigem.TableDefs(strTabela)

'' SELECIONA OS REGISTROS DA ORIGEM
Dim rstOrigem As DAO.Recordset
Set rstOrigem = dbOrigem.OpenRecordset("Select * from " & strTabela & "")


''==============================''
''           DESTINO
''==============================''

'' POSICIONA O BANCO DE DESTINO
Dim dbDestino As DAO.Database
Set dbDestino = OpenDatabase(strDestino, False, False, "MS Access;PWD=" & SenhaBanco)

'' SELECIONA A TABELA DE DESTINO
Dim tdfNew As DAO.TableDef
Set tdfNew = dbDestino.CreateTableDef(strTabela)

'' SELECIONA OS REGISTROS DA ORIGEM
Dim rstDestino As DAO.Recordset
Set rstDestino = dbDestino.OpenRecordset("Select * from " & strTabela & "")

Dim x As Integer

'Saida Now(), "ExportarDadosTabelaOrigemParaTabelaDestino.log"

While Not rstOrigem.EOF

    rstDestino.AddNew

    For x = 0 To dbOrigem.TableDefs(strTabela).Fields.Count - 1

        With tdfNew
             rstDestino.Fields(dbDestino.TableDefs(strTabela).Fields(x).Properties("name")) = rstOrigem.Fields(dbOrigem.TableDefs(strTabela).Fields(x).Properties("name"))
        End With

    Next x

    rstDestino.Update
    rstOrigem.MoveNext

Wend

'Saida Now(), "ExportarDadosTabelaOrigemParaTabelaDestino.log"

rstDestino.Close
rstOrigem.Close
dbDestino.Close
dbOrigem.Close

Set rstDestino = Nothing
Set rstOrigem = Nothing
Set dbDestino = Nothing
Set dbOrigem = Nothing

End Sub


Sub ReceberAtualizacoes(ByVal strPastaChegada As String)
''' RECEBER ATUALIZAÇÕES
Dim strBancoOrigem As String: strBancoOrigem = Range(BancoLocal)
Dim Matriz As Variant
Dim x As Long
Dim y As Long
Dim lista As String
Dim CATEGORIA As String

    Matriz = Array()

    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ''' LISTAR ARQUIVOS COMPACTADOS, DESCOMPACTA-LOS E EXCLUI-LOS
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

    ''' LISTAR
    lista = ListarDiretorio(strPastaChegada, "*.zip")
    Matriz = Split(lista, ";")

    For x = 0 To UBound(Matriz)
        ''' DESCOMPACTAR
'        UnZip strPastaChegada & Matriz(x), strPastaChegada
        DesCompact strPastaChegada & Matriz(x), strPastaChegada

        ''' EXCLUIR
        If Dir$(strPastaChegada & Matriz(x)) <> "" Then Kill strPastaChegada & Matriz(x)
    Next x

    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ''' LISTAR ARQUIVOS DESCOMPACTADOS, EXPORTA-LOS AO BANCO PRINCIPAL E EXCLUI-LOS
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

    ''' LISTAR
    lista = ListarDiretorio(strPastaChegada, "*.mdb")
    Matriz = Split(lista, ";")

    For y = 0 To UBound(Matriz)

        CATEGORIA = getFileStep(CStr(Matriz(y)))

        Select Case CATEGORIA

            Case "TRANSITO"

                BancoEmTransito_Importar strPastaChegada & Matriz(y)

            Case "ATUALIZACAO"

                ''' LIMPAR CONFIGURAÇÕES ANTIGAS
                admCategoriaLimparTabela strBancoOrigem

                ''' CARREGAR NOVAS CONFIGURAÇÕES
                ExportarDadosTabelaOrigemParaTabelaDestino strPastaChegada & Matriz(y), strBancoOrigem, "admCategorias"

                ''' ATUALIZAR GUIAS DE APOIO
                admAtualizarGuiaDeApoio


        End Select

        ''' EXCLUIR
        If Dir$(strPastaChegada & Matriz(y)) <> "" Then Kill strPastaChegada & Matriz(y)
    Next y


End Sub

Sub BancoEmTransito_Importar(ByVal strBancoEmTransito As String)
On Error GoTo cmdImportar_err

Dim strBancoLocal As String: strBancoLocal = Range(BancoLocal)

Dim dbOrigem As DAO.Database
Dim rstOrcamentoORIGEM As DAO.Recordset

Set dbOrigem = DBEngine.OpenDatabase(strBancoEmTransito, False, False, "MS Access;PWD=" & SenhaBanco)
Set rstOrcamentoORIGEM = dbOrigem.OpenRecordset("Select * from Orcamentos Order by Vendedor,Controle")

    While Not rstOrcamentoORIGEM.EOF
        ''' IMPORTAR ORÇAMENTOS EM TRANSITO
        ExportarOrcamento strBancoEmTransito, strBancoLocal, rstOrcamentoORIGEM.Fields("Controle"), rstOrcamentoORIGEM.Fields("Vendedor")
        rstOrcamentoORIGEM.MoveNext

    Wend


cmdImportar_Fim:

    MsgBox "Importação concluída", vbInformation + vbOKOnly, "Importar de Orçamento(s)"

    Exit Sub
cmdImportar_err:
    MsgBox Err.Description, vbCritical + vbOKOnly, "Importar de Orçamento(s)"
    Resume cmdImportar_Fim

End Sub

Function CriarArquivoDeAtualizacaoDoSistema() As String
Dim strBancoOrigem As String: strBancoOrigem = Range(BancoLocal)
Dim strBancoDestino As String: strBancoDestino = pathWorkSheetAddress & Controle & "_db" & "ATUALIZACAO" & ".mdb"
Dim strArquivoCompactado As String: strArquivoCompactado = Left(strBancoDestino, Len(strBancoDestino) - 3) & "zip"

    ''' CRIA BASE DE DADOS PARA EXPORTAÇÃO DE DADOS
    CriarBancoParaExportacao strBancoDestino

    ''' CRIAR TABELA(S) EM BASE DE DADOS DE EXPORTAÇÃO
    CriarTabelaEmBancoParaExportacao strBancoOrigem, strBancoDestino, "admCategorias"

    ''' EXPORTAR DADOS DA TABELA ADMINISTRATIVA DO SISTEMA
    ExportarDadosTabelaOrigemParaTabelaDestino strBancoOrigem, strBancoDestino, "admCategorias"

    ''' COMPACTA BASE DE DADOS
    Zip strBancoDestino, strArquivoCompactado

    ''' DELETA BASE DE DADOS TEMPORARIA
    If Dir$(strBancoDestino) <> "" Then Kill strBancoDestino

    ''' RETORNO DE NOME DO ARQUIVO CRIADO
    CriarArquivoDeAtualizacaoDoSistema = strArquivoCompactado

End Function

