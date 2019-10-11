Attribute VB_Name = "mod_03_Exportar"
Option Explicit

Public Sub ExportarOrcamento(Origem As String, _
                            Destino As String, _
                            strControle As String, _
                            strVendedor As String)
                            
On Error GoTo ExportarOrcamento_err

'Dim ORIGEM As String: ORIGEM = Range(BancoLocal)
'Dim DESTINO As String: DESTINO = pathWorkSheetAddress & "dbVendedor.mdb"

'   BANCO DE DADOS
Dim dbOrigem As DAO.Database
Set dbOrigem = DBEngine.OpenDatabase(Origem, False, False, "MS Access;PWD=" & SenhaBanco)

Dim dbDestino As DAO.Database
Set dbDestino = DBEngine.OpenDatabase(Destino, False, False, "MS Access;PWD=" & SenhaBanco)

Dim rstOrcamentoORIGEM As DAO.Recordset
Set rstOrcamentoORIGEM = dbOrigem.OpenRecordset("Select * from Orcamentos where controle = '" & strControle & "' and Vendedor = '" & strVendedor & "'")

Dim rstOrcamentoDESTINO As DAO.Recordset
Set rstOrcamentoDESTINO = dbDestino.OpenRecordset("Select * from Orcamentos where controle = '" & strControle & "' and Vendedor = '" & strVendedor & "'")

Dim x As Integer

While Not rstOrcamentoORIGEM.EOF

       
        If Not rstOrcamentoDESTINO.EOF Then
            rstOrcamentoDESTINO.Edit
        Else
            rstOrcamentoDESTINO.AddNew
            rstOrcamentoDESTINO.Fields("VENDEDOR") = rstOrcamentoORIGEM.Fields("VENDEDOR")
            rstOrcamentoDESTINO.Fields("CONTROLE") = rstOrcamentoORIGEM.Fields("CONTROLE")
        End If
        
        rstOrcamentoDESTINO.Fields("COD_VENDEDOR") = rstOrcamentoORIGEM.Fields("COD_VENDEDOR")
        
        rstOrcamentoDESTINO.Fields("ID_ETAPA") = rstOrcamentoORIGEM.Fields("ID_ETAPA")
        rstOrcamentoDESTINO.Fields("STATUS") = rstOrcamentoORIGEM.Fields("STATUS")
        rstOrcamentoDESTINO.Fields("DEPARTAMENTO") = rstOrcamentoORIGEM.Fields("DEPARTAMENTO")
        
        rstOrcamentoDESTINO.Fields("CLIENTE") = rstOrcamentoORIGEM.Fields("CLIENTE")
        rstOrcamentoDESTINO.Fields("RESPONSAVEL") = rstOrcamentoORIGEM.Fields("RESPONSAVEL")
        rstOrcamentoDESTINO.Fields("PROJETO") = IIf((rstOrcamentoORIGEM.Fields("PROJETO")) <> "", rstOrcamentoORIGEM.Fields("PROJETO"), " ")
        rstOrcamentoDESTINO.Fields("LINHA_PRODUTO") = rstOrcamentoORIGEM.Fields("LINHA_PRODUTO")
        rstOrcamentoDESTINO.Fields("DT_PEDIDO") = rstOrcamentoORIGEM.Fields("DT_PEDIDO")
        rstOrcamentoDESTINO.Fields("PREV_ENTREGA") = rstOrcamentoORIGEM.Fields("PREV_ENTREGA")
        rstOrcamentoDESTINO.Fields("VALOR_PROJETO") = rstOrcamentoORIGEM.Fields("VALOR_PROJETO")
        rstOrcamentoDESTINO.Fields("PUBLISHER") = rstOrcamentoORIGEM.Fields("PUBLISHER")
        rstOrcamentoDESTINO.Fields("JOURNAL") = rstOrcamentoORIGEM.Fields("JOURNAL")
        rstOrcamentoDESTINO.Fields("PAGS") = rstOrcamentoORIGEM.Fields("PAGS")

        'FECHADO COM CLIENTE
        For x = 1 To 8
            rstOrcamentoDESTINO.Fields(x & "_FECHADO") = rstOrcamentoORIGEM.Fields(x & "_FECHADO")
        Next x

        'VENDA
        For x = 1 To 8
            rstOrcamentoDESTINO.Fields(x & "_VENDA") = rstOrcamentoORIGEM.Fields(x & "_VENDA")
        Next x

        'IDIOMA
        For x = 1 To 8
            rstOrcamentoDESTINO.Fields(x & "_IDIOMA") = rstOrcamentoORIGEM.Fields(x & "_IDIOMA")
        Next x

        'TIRAGEM
        For x = 1 To 8
            rstOrcamentoDESTINO.Fields(x & "_TIRAGEM") = rstOrcamentoORIGEM.Fields(x & "_TIRAGEM")
        Next x

        'ESPECIFICACAO
        For x = 1 To 8
            rstOrcamentoDESTINO.Fields(x & "_ESPECIFICACAO") = rstOrcamentoORIGEM.Fields(x & "_ESPECIFICACAO")
        Next x

        'MOEDA
        For x = 1 To 8
            rstOrcamentoDESTINO.Fields(x & "_MOEDA") = rstOrcamentoORIGEM.Fields(x & "_MOEDA")
        Next x

        'ROYALTY PERCENTUAL
        For x = 1 To 8
            rstOrcamentoDESTINO.Fields(x & "_ROYALTY_PERCENTUAL") = rstOrcamentoORIGEM.Fields(x & "_ROYALTY_PERCENTUAL")
        Next x

        'ROYALTY ESPECIE
        For x = 1 To 8
            rstOrcamentoDESTINO.Fields(x & "_ROYALTY_ESPECIE") = rstOrcamentoORIGEM.Fields(x & "_ROYALTY_ESPECIE")
        Next x

        'RE IMPRESSAO
        For x = 1 To 8
            rstOrcamentoDESTINO.Fields(x & "_RE_IMPRESSAO") = rstOrcamentoORIGEM.Fields(x & "_RE_IMPRESSAO")
        Next x

        'DESCONTO
        For x = 1 To 8
            rstOrcamentoDESTINO.Fields(x & "_DESCONTO") = rstOrcamentoORIGEM.Fields(x & "_DESCONTO")
        Next x

        'TIPO
        For x = 1 To 4
            rstOrcamentoDESTINO.Fields(x & "_TIPO") = rstOrcamentoORIGEM.Fields(x & "_TIPO")
        Next x

        'PAPEL
        For x = 1 To 4
            rstOrcamentoDESTINO.Fields(x & "_PAPEL") = rstOrcamentoORIGEM.Fields(x & "_PAPEL")
        Next x

        'PAGINAS
        For x = 1 To 4
            rstOrcamentoDESTINO.Fields(x & "_PAGINAS") = rstOrcamentoORIGEM.Fields(x & "_PAGINAS")
        Next x

        'IMPRESSAO
        For x = 1 To 4
            rstOrcamentoDESTINO.Fields(x & "_IMPRESSAO") = rstOrcamentoORIGEM.Fields(x & "_IMPRESSAO")
        Next x

        'FORMATO
        For x = 1 To 4
            rstOrcamentoDESTINO.Fields(x & "_FORMATO") = rstOrcamentoORIGEM.Fields(x & "_FORMATO")
        Next x

        'ACABAMENTO
        For x = 1 To 4
            If rstOrcamentoORIGEM.Fields(x & "_ACABAMENTO") <> "" Then rstOrcamentoDESTINO.Fields(x & "_ACABAMENTO") = rstOrcamentoORIGEM.Fields(x & "_ACABAMENTO")
        Next x
    
        rstOrcamentoDESTINO.Update
        rstOrcamentoORIGEM.MoveNext
        
Wend
    
ExportarCusto Origem, Destino, strControle, strVendedor
    
ExportarOrcamento_Fim:
    rstOrcamentoORIGEM.Close
    rstOrcamentoDESTINO.Close
    dbOrigem.Close
    dbDestino.Close
    
    Set dbOrigem = Nothing
    Set rstOrcamentoORIGEM = Nothing
    Set rstOrcamentoDESTINO = Nothing
    
'     MsgBox "ok!", , "Exportar Orçamento"
    
    Exit Sub
ExportarOrcamento_err:
    MsgBox Err.Description, vbCritical + vbOKOnly, "Exportar Orçamento"
    Resume ExportarOrcamento_Fim


End Sub

Private Sub ExportarCusto(Origem As String, _
                            Destino As String, _
                            strControle As String, _
                            strVendedor As String)
                            
On Error GoTo ExportarCusto_err

'Dim ORIGEM As String: ORIGEM = Range(BancoLocal)
'Dim DESTINO As String: DESTINO = pathWorkSheetAddress & "dbVendedor.mdb"

'   BANCO DE DADOS
Dim dbOrigem As DAO.Database
Set dbOrigem = DBEngine.OpenDatabase(Origem, False, False, "MS Access;PWD=" & SenhaBanco)

Dim dbDestino As DAO.Database
Set dbDestino = DBEngine.OpenDatabase(Destino, False, False, "MS Access;PWD=" & SenhaBanco)

Dim rstCustosORIGEM As DAO.Recordset
Set rstCustosORIGEM = dbOrigem.OpenRecordset("Select * from OrcamentosCustos where controle = '" & strControle & "' and Vendedor = '" & strVendedor & "'")

Dim rstCustosDESTINO As DAO.Recordset
Set rstCustosDESTINO = dbDestino.OpenRecordset("Select * from OrcamentosCustos where controle = '" & strControle & "' and Vendedor = '" & strVendedor & "'")

Dim x As Integer

While Not rstCustosORIGEM.EOF
        
        If Not rstCustosDESTINO.EOF Then
            rstCustosDESTINO.Edit
        Else
            rstCustosDESTINO.AddNew
            rstCustosDESTINO.Fields("VENDEDOR") = rstCustosORIGEM.Fields("VENDEDOR")
            rstCustosDESTINO.Fields("CONTROLE") = rstCustosORIGEM.Fields("CONTROLE")
        End If

    
        'INDEXACAO
        For x = 1 To 8
            rstCustosDESTINO.Fields(x & "_INDEXACAO") = rstCustosORIGEM.Fields(x & "_INDEXACAO")
        Next x
        
        'TRADUCAO
        For x = 1 To 8
            rstCustosDESTINO.Fields(x & "_TRADUCAO") = rstCustosORIGEM.Fields(x & "_TRADUCAO")
        Next x
        
        'REVISAO ORTOGRAFICA
        For x = 1 To 8
            rstCustosDESTINO.Fields(x & "_REVISAO_ORTOGRAFICA") = rstCustosORIGEM.Fields(x & "_REVISAO_ORTOGRAFICA")
        Next x
        
        'REVISAO MEDICA
        For x = 1 To 8
            rstCustosDESTINO.Fields(x & "_REVISAO_MEDICA") = rstCustosORIGEM.Fields(x & "_REVISAO_MEDICA")
        Next x
        
        'CRIACAO
        For x = 1 To 8
            rstCustosDESTINO.Fields(x & "_CRIACAO") = rstCustosORIGEM.Fields(x & "_CRIACAO")
        Next x
        
        'ILUSTRACAO
        For x = 1 To 8
            rstCustosDESTINO.Fields(x & "_ILUSTRACAO") = rstCustosORIGEM.Fields(x & "_ILUSTRACAO")
        Next x
        
        'REVISAO
        For x = 1 To 8
            rstCustosDESTINO.Fields(x & "_REVISAO") = rstCustosORIGEM.Fields(x & "_REVISAO")
        Next x
        
        'DIAGRAMACAO
        For x = 1 To 8
            rstCustosDESTINO.Fields(x & "_DIAGRAMACAO") = rstCustosORIGEM.Fields(x & "_DIAGRAMACAO")
        Next x
        
        'MEDICO
        For x = 1 To 8
            rstCustosDESTINO.Fields(x & "_MEDICO") = rstCustosORIGEM.Fields(x & "_MEDICO")
        Next x
        
        'GRAFICA
        For x = 1 To 8
            rstCustosDESTINO.Fields(x & "_GRAFICA") = rstCustosORIGEM.Fields(x & "_GRAFICA")
        Next x
    
        'MIDIA
        For x = 1 To 8
            rstCustosDESTINO.Fields(x & "_MIDIA") = rstCustosORIGEM.Fields(x & "_MIDIA")
        Next x
    
        'CORREIO
        For x = 1 To 8
            rstCustosDESTINO.Fields(x & "_CORREIO") = rstCustosORIGEM.Fields(x & "_CORREIO")
        Next x
    
        'ULTIMA CAPA
        For x = 1 To 8
            rstCustosDESTINO.Fields(x & "_ULTIMA_CAPA") = rstCustosORIGEM.Fields(x & "_ULTIMA_CAPA")
        Next x
    
        'IMPORT
        For x = 1 To 8
            rstCustosDESTINO.Fields(x & "_IMPORT") = rstCustosORIGEM.Fields(x & "_IMPORT")
        Next x
    
        'TRANSPORTE NACIONAL
        For x = 1 To 8
            rstCustosDESTINO.Fields(x & "_TRANSPORTE_NACIONAL") = rstCustosORIGEM.Fields(x & "_TRANSPORTE_NACIONAL")
        Next x
    
        'TRANSPORTE_INTERNACIONAL
        For x = 1 To 8
            rstCustosDESTINO.Fields(x & "_TRANSPORTE_INTERNACIONAL") = rstCustosORIGEM.Fields(x & "_TRANSPORTE_INTERNACIONAL")
        Next x
    
        'SEGUROS
        For x = 1 To 8
            rstCustosDESTINO.Fields(x & "_SEGUROS") = rstCustosORIGEM.Fields(x & "_SEGUROS")
        Next x
    
        'EXTRAS
        For x = 1 To 8
            rstCustosDESTINO.Fields(x & "_EXTRAS") = rstCustosORIGEM.Fields(x & "_EXTRAS")
        Next x
    
        'EDITOR_FEE
        For x = 1 To 8
            rstCustosDESTINO.Fields(x & "_EDITOR_FEE") = rstCustosORIGEM.Fields(x & "_EDITOR_FEE")
        Next x
    
        'DESP_VIAGEM
        For x = 1 To 8
            rstCustosDESTINO.Fields(x & "_DESP_VIAGEM") = rstCustosORIGEM.Fields(x & "_DESP_VIAGEM")
        Next x
    
        'OUTROS
        For x = 1 To 8
            rstCustosDESTINO.Fields(x & "_OUTROS") = rstCustosORIGEM.Fields(x & "_OUTROS")
        Next x
    
        rstCustosDESTINO.Update
        rstCustosORIGEM.MoveNext
        
Wend

ExportarAnexos Origem, Destino, strControle, strVendedor
    
ExportarCusto_Fim:
    rstCustosORIGEM.Close
    rstCustosDESTINO.Close
    dbOrigem.Close
    dbDestino.Close
    
    Set dbOrigem = Nothing
    Set rstCustosORIGEM = Nothing
    Set rstCustosDESTINO = Nothing
    
'    MsgBox "ok!", , "Exportar Custos"
    
    Exit Sub
ExportarCusto_err:
    MsgBox Err.Description, vbCritical + vbOKOnly, "Exportar Custos"
    Resume ExportarCusto_Fim


End Sub

Private Sub ExportarAnexos(Origem As String, _
                            Destino As String, _
                            strControle As String, _
                            strVendedor As String)

On Error GoTo ExportarAnexos_err

'   BANCO DE DADOS
Dim dbOrigem As DAO.Database
Set dbOrigem = DBEngine.OpenDatabase(Origem, False, False, "MS Access;PWD=" & SenhaBanco)

Dim dbDestino As DAO.Database
Set dbDestino = DBEngine.OpenDatabase(Destino, False, False, "MS Access;PWD=" & SenhaBanco)

Dim rstAnexosORIGEM As DAO.Recordset
Set rstAnexosORIGEM = dbOrigem.OpenRecordset("Select * from OrcamentosAnexos where controle = '" & strControle & "' and Vendedor = '" & strVendedor & "'")

Dim rstAnexosDESTINO As DAO.Recordset
Set rstAnexosDESTINO = dbDestino.OpenRecordset("Select * from OrcamentosAnexos where controle = '" & strControle & "' and Vendedor = '" & strVendedor & "'")

Dim x As Integer

While Not rstAnexosORIGEM.EOF
        
        If Not rstAnexosDESTINO.EOF Then
            rstAnexosDESTINO.Edit
        Else
            rstAnexosDESTINO.AddNew
            rstAnexosDESTINO.Fields("VENDEDOR") = rstAnexosORIGEM.Fields("VENDEDOR")
            rstAnexosDESTINO.Fields("CONTROLE") = rstAnexosORIGEM.Fields("CONTROLE")
        End If

        rstAnexosDESTINO.Fields("PROPRIEDADE") = rstAnexosORIGEM.Fields("PROPRIEDADE")
        rstAnexosDESTINO.Fields("DESCRICAO") = rstAnexosORIGEM.Fields("DESCRICAO")
        rstAnexosDESTINO.Fields("VALOR_01") = rstAnexosORIGEM.Fields("VALOR_01")
        rstAnexosDESTINO.Fields("VALOR_02") = rstAnexosORIGEM.Fields("VALOR_02")
    
        rstAnexosDESTINO.Update
        
        If Not rstAnexosDESTINO.EOF Then
            rstAnexosDESTINO.MoveNext
        End If
        
        rstAnexosORIGEM.MoveNext
        
Wend
    
ExportarAnexos_Fim:
    rstAnexosORIGEM.Close
    rstAnexosDESTINO.Close
    dbOrigem.Close
    dbDestino.Close
    
    Set dbOrigem = Nothing
    Set rstAnexosORIGEM = Nothing
    Set rstAnexosDESTINO = Nothing
    
'    MsgBox "ok!", , "Exportar Anexos"
    
    Exit Sub
ExportarAnexos_err:
    MsgBox Err.Description, vbCritical + vbOKOnly, "Exportar Anexos"
    Resume ExportarAnexos_Fim


End Sub

