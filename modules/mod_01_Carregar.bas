Attribute VB_Name = "mod_01_Carregar"
Option Base 1
Option Explicit

Public Function carregarOrcamento( _
                                    BaseDeDados As String, _
                                    strControle As String, _
                                    strVendedor As String)
On Error GoTo CarregarOrcamento_err

Dim dbOrcamento As DAO.Database
Dim rstCarregarOrcamento As DAO.Recordset
Dim rstCarregarCustos As DAO.Recordset

Dim L As Integer, c As Integer ' L = LINHA | C = COLUNA
Dim x As Integer ' contador de linhas

Set dbOrcamento = DBEngine.OpenDatabase(BaseDeDados, False, False, "MS Access;PWD=" & SenhaBanco)
Set rstCarregarOrcamento = dbOrcamento.OpenRecordset("Select * from Orcamentos where controle = '" & strControle & "' and Vendedor = '" & strVendedor & "'")


    Application.ScreenUpdating = False
    DesbloqueioDeGuia SenhaBloqueio
    
    '#########################
    '   ORÇAMENTO
    '#########################
    
    With rstCarregarOrcamento
    
        Range("C3").value = .Fields("VENDEDOR")
        
        Range("C4").value = .Fields("CLIENTE")
        Range("C5").value = .Fields("RESPONSAVEL")
        
        Range("C6").value = .Fields("PROJETO")
        Range("G3").value = .Fields("DT_PEDIDO")
        Range("G4").value = .Fields("PREV_ENTREGA")
        Range("J4").value = .Fields("VALOR_PROJETO")
        Range("J3").value = .Fields("STATUS")
        Range("C8").value = .Fields("PUBLISHER")
        Range("C9").value = .Fields("JOURNAL")
        Range("C10").value = .Fields("PAGS")
    
        'FECHADO COM CLIENTE
        L = 12
        c = 3
        For x = 1 To 8
            Cells(L, c).value = .Fields(x & "_FECHADO")
            c = c + 1
        Next x
        
        'LINHA
        L = 13
        c = 3
        For x = 1 To 4
            Cells(L, c).value = .Fields(x & "_LINHA_PRODUTO")
            c = c + 1
        Next x
        
        'FASCICULOS
        L = 14
        c = 3
        For x = 1 To 4
            Cells(L, c).value = .Fields(x & "_FASCICULOS")
            c = c + 1
        Next x
        
        'VENDA
        L = 15
        c = 3
        For x = 1 To 8
            Cells(L, c).value = .Fields(x & "_VENDA")
            c = c + 1
        Next x
        
        ''''''''''''''''''''''''''''''''''''' [IMPOSTOS]
        ''''''''''''''''''''''''''''''''''''' [IMPOSTOS]
        ''''''''''''''''''''''''''''''''''''' [IMPOSTOS]
        ''''''''''''''''''''''''''''''''''''' [IMPOSTOS]
    
        'IDIOMA
        L = 17
        c = 3
        For x = 1 To 8
            Cells(L, c).value = .Fields(x & "_IDIOMA")
            c = c + 1
        Next x
    
        'TIRAGEM
        L = 18
        c = 3
        For x = 1 To 8
            Cells(L, c).value = .Fields(x & "_TIRAGEM")
            c = c + 1
        Next x
    
        'ESPECIFICACAO
        L = 19
        c = 3
        For x = 1 To 8
            Cells(L, c).value = .Fields(x & "_ESPECIFICACAO")
            c = c + 1
        Next x
    
        'MOEDA
        L = 20
        c = 3
        For x = 1 To 8
            Cells(L, c).value = .Fields(x & "_MOEDA")
            c = c + 1
        Next x
    
        'ROYALTY PERCENTUAL
        L = 21
        c = 3
        For x = 1 To 8
            Cells(L, c).value = .Fields(x & "_ROYALTY_PERCENTUAL")
            c = c + 1
        Next x
    
        'ROYALTY ESPECIE
        L = 22
        c = 3
        For x = 1 To 8
            Cells(L, c).value = .Fields(x & "_ROYALTY_ESPECIE")
            c = c + 1
        Next x
    
        'RE IMPRESSAO
        L = 23
        c = 3
        For x = 1 To 8
            Cells(L, c).value = .Fields(x & "_RE_IMPRESSAO")
            c = c + 1
        Next x
    
        'TIPO
        L = 25
        c = 3
        For x = 1 To 4
            Cells(L, c).value = .Fields(x & "_TIPO")
            c = c + 1
        Next x
    
        'PAPEL
        L = 26
        c = 3
        For x = 1 To 4
            Cells(L, c).value = .Fields(x & "_PAPEL")
            c = c + 1
        Next x
    
        'PAGINAS
        L = 27
        c = 3
        For x = 1 To 4
            Cells(L, c).value = .Fields(x & "_PAGINAS")
            c = c + 1
        Next x
    
        'IMPRESSAO
        L = 28
        c = 3
        For x = 1 To 4
            Cells(L, c).value = .Fields(x & "_IMPRESSAO")
            c = c + 1
        Next x
        
        'FORMATO
        L = 29
        c = 3
        For x = 1 To 4
            Cells(L, c).value = .Fields(x & "_FORMATO")
            c = c + 1
        Next x
    
        'ACABAMENTO
        L = 31
        c = 2
        For x = 1 To 4
            Cells(L, c).value = .Fields(x & "_ACABAMENTO")
            L = L + 1
        Next x
                
'        'DESCONTO
'        L = 61
'        c = 3
'        For x = 1 To 8
'            Cells(L, c).Value = .Fields(x & "_DESCONTO")
'            c = c + 1
'        Next x
        
        'PREÇO MKT
        L = 65
        c = 3
        For x = 1 To 4
            Cells(L, c).value = .Fields(x & "_PrecoMKT")
            c = c + 1
        Next x
        
        'DESCONTO PADRÃO
        L = 71
        c = 3
        For x = 1 To 4
            Cells(L, c).value = .Fields(x & "_DescontoPadrao")
            c = c + 1
        Next x
        
        'PREÇO COMPRA TOTAL
        L = 73
        c = 3
        For x = 1 To 4
            Cells(L, c).value = .Fields(x & "_PrecoTotal")
            c = c + 1
        Next x
        
        'ARREDONDAMENTO
        L = 83
        c = 3
        For x = 1 To 4
            Cells(L, c).value = .Fields(x & "_Arredondamento")
            c = c + 1
        Next x
        
    
    End With
    
    '#########################
    '   CUSTOS
    '#########################
    
carregar_custos:
    
    Set rstCarregarCustos = dbOrcamento.OpenRecordset("Select * from OrcamentosCustos where controle = '" & strControle & "' and Vendedor = '" & strVendedor & "'")
    
    If Not rstCarregarCustos.EOF Then
        
        With rstCarregarCustos
        
            'INDEXACAO
            L = 37
            c = 3
            For x = 1 To 8
                Cells(L, c).value = .Fields(x & "_INDEXACAO")
                c = c + 1
            Next x
            
            'TRADUCAO
            L = 38
            c = 3
            For x = 1 To 8
                Cells(L, c).value = .Fields(x & "_TRADUCAO")
                c = c + 1
            Next x
            
            'REVISAO ORTOGRAFICA
            L = 39
            c = 3
            For x = 1 To 8
                Cells(L, c).value = .Fields(x & "_REVISAO_ORTOGRAFICA")
                c = c + 1
            Next x
            
            'REVISAO MEDICA
            L = 40
            c = 3
            For x = 1 To 8
                Cells(L, c).value = .Fields(x & "_REVISAO_MEDICA")
                c = c + 1
            Next x
            
            'CRIACAO
            L = 41
            c = 3
            For x = 1 To 8
                Cells(L, c).value = .Fields(x & "_CRIACAO")
                c = c + 1
            Next x
            
            'ILUSTRACAO
            L = 42
            c = 3
            For x = 1 To 8
                Cells(L, c).value = .Fields(x & "_ILUSTRACAO")
                c = c + 1
            Next x
            
            'REVISAO
            L = 43
            c = 3
            For x = 1 To 8
                Cells(L, c).value = .Fields(x & "_REVISAO")
                c = c + 1
            Next x
            
            'DIAGRAMACAO
            L = 44
            c = 3
            For x = 1 To 8
                Cells(L, c).value = .Fields(x & "_DIAGRAMACAO")
                c = c + 1
            Next x
            
            'MEDICO
            L = 45
            c = 3
            For x = 1 To 8
                Cells(L, c).value = .Fields(x & "_MEDICO")
                c = c + 1
            Next x
            
            'GRAFICA
            L = 46
            c = 3
            For x = 1 To 8
                Cells(L, c).value = .Fields(x & "_GRAFICA")
                c = c + 1
            Next x
            
            'MIDIA
            L = 47
            c = 3
            For x = 1 To 8
                Cells(L, c).value = .Fields(x & "_MIDIA")
                c = c + 1
            Next x
        
            'CORREIO
            L = 48
            c = 3
            For x = 1 To 8
                Cells(L, c).value = .Fields(x & "_CORREIO")
                c = c + 1
            Next x
        
        
            'ULTIMA CAPA
            L = 49
            c = 3
            For x = 1 To 8
                Cells(L, c).value = .Fields(x & "_ULTIMA_CAPA")
                c = c + 1
            Next x
        
            'IMPORT
            L = 50
            c = 3
            For x = 1 To 8
                Cells(L, c).value = .Fields(x & "_IMPORT")
                c = c + 1
            Next x
        
            'TRANSPORTE NACIONAL
            L = 51
            c = 3
            For x = 1 To 8
                Cells(L, c).value = .Fields(x & "_TRANSPORTE_NACIONAL")
                c = c + 1
            Next x
        
            'TRANSPORTE INTERNACIONAL
            L = 52
            c = 3
            For x = 1 To 8
                Cells(L, c).value = .Fields(x & "_TRANSPORTE_INTERNACIONAL")
                c = c + 1
            Next x
        
            'SEGUROS
            L = 53
            c = 3
            For x = 1 To 8
                Cells(L, c).value = .Fields(x & "_SEGUROS")
                c = c + 1
            Next x
        
            'EXTRAS
            L = 54
            c = 3
            For x = 1 To 8
                Cells(L, c).value = .Fields(x & "_EXTRAS")
                c = c + 1
            Next x
        
            'EDITOR FEE
            L = 55
            c = 3
            For x = 1 To 8
                Cells(L, c).value = .Fields(x & "_EDITOR_FEE")
                c = c + 1
            Next x
        
            'DESP VIAGEM
            L = 56
            c = 3
            For x = 1 To 8
                Cells(L, c).value = .Fields(x & "_DESP_VIAGEM")
                c = c + 1
            Next x
        
            'OUTROS
            L = 57
            c = 3
            For x = 1 To 8
                Cells(L, c).value = .Fields(x & "_OUTROS")
                c = c + 1
            Next x
            
        End With
    
    Else
    
        admOrcamentoNovoCustosProducao BaseDeDados, strControle, strVendedor
        
        GoTo carregar_custos
        
    
    End If
    
    CarregarAnexoLinha BaseDeDados, strControle, strVendedor, 3, 12
    CarregarAnexoMoeda BaseDeDados, strControle, strVendedor, 3, 16
    CarregarAnexoVenda BaseDeDados, strControle, strVendedor, 3, 19
    CarregarAnexoDesconto BaseDeDados, strControle, strVendedor, 3, 22
    
'    CarregarAnexoTraducao BaseDeDados, strControle, strVendedor, 3, 29
'    CarregarAnexoRevisao BaseDeDados, strControle, strVendedor, 3, 32
'    CarregarAnexoDiagramacao BaseDeDados, strControle, strVendedor, 3, 35
    
    
'    CarregarAnexoArquivo BaseDeDados, strControle, strVendedor, CInt(Right(ArquivoInicio, Len(ArquivoInicio) - 1)), 2
    
    BloqueioDeGuia SenhaBloqueio
    Application.ScreenUpdating = True


CarregarOrcamento_Fim:
    rstCarregarOrcamento.Close
    rstCarregarCustos.Close
    dbOrcamento.Close
    
    Set dbOrcamento = Nothing
    Set rstCarregarCustos = Nothing
    Set rstCarregarOrcamento = Nothing
    
    Exit Function
CarregarOrcamento_err:
    MsgBox Err.Description
    Resume CarregarOrcamento_Fim


End Function

Public Function CarregarAnexoDesconto( _
                                    BaseDeDados As String, _
                                    strControle As String, _
                                    strVendedor As String, _
                                    intLinha As Integer, _
                                    intColuna As Integer)
                                    
On Error GoTo CarregarAnexoDesconto_err
                                    
Dim dbOrcamento As DAO.Database
Dim rstCarregarAnexoDesconto As DAO.Recordset

'Dim L As Integer, c As Integer ' L = LINHA | C = COLUNA
'Dim x, y As Integer ' contador de linhas

Set dbOrcamento = DBEngine.OpenDatabase(BaseDeDados, False, False, "MS Access;PWD=" & SenhaBanco)
Set rstCarregarAnexoDesconto = dbOrcamento.OpenRecordset("Select * from OrcamentosAnexos where controle = '" & strControle & "' and Vendedor = '" & strVendedor & "' and PROPRIEDADE = 'DESCONTO'")
    
rstCarregarAnexoDesconto.MoveLast
rstCarregarAnexoDesconto.MoveFirst
'y = rstCarregarAnexoDesconto.RecordCount
    
'    l = 3
'    c = 22

While Not rstCarregarAnexoDesconto.EOF
    
'    For x = 1 To y
        
        With rstCarregarAnexoDesconto
        
            Cells(intLinha, intColuna + 1).value = .Fields("DESCRICAO")
            Cells(intLinha, intColuna).value = Val(.Fields("VALOR_01"))
            rstCarregarAnexoDesconto.MoveNext
            
        End With
        
        intLinha = intLinha + 1
'    Next x

Wend

CarregarAnexoDesconto_Fim:
    rstCarregarAnexoDesconto.Close
    dbOrcamento.Close
    
    Set dbOrcamento = Nothing
    Set rstCarregarAnexoDesconto = Nothing
    
    Exit Function
CarregarAnexoDesconto_err:
    If Err.Number = "3021" Then
        MsgBox "ATENÇÃO: Não exitem registros de Desconto", vbInformation + vbOKOnly, "Registros de Desconto"
    Else
        MsgBox Err.Number & Space(2) & Err.Description, , "Registros de Desconto"
    End If
    
    Resume CarregarAnexoDesconto_Fim

End Function


Public Function CarregarAnexoTraducao( _
                                    BaseDeDados As String, _
                                    strControle As String, _
                                    strVendedor As String, _
                                    intLinha As Integer, _
                                    intColuna As Integer)
                                    
On Error GoTo CarregarAnexoTraducao_err
                                    
Dim dbOrcamento As DAO.Database
Dim rstCarregarAnexoTraducao As DAO.Recordset

'Dim L As Integer, c As Integer ' L = LINHA | C = COLUNA
'Dim x, y As Integer ' contador de linhas

Set dbOrcamento = DBEngine.OpenDatabase(BaseDeDados, False, False, "MS Access;PWD=" & SenhaBanco)
Set rstCarregarAnexoTraducao = dbOrcamento.OpenRecordset("Select * from OrcamentosAnexos where controle = '" & strControle & "' and Vendedor = '" & strVendedor & "' and PROPRIEDADE = 'TRADUÇÃO'")
    
rstCarregarAnexoTraducao.MoveLast
rstCarregarAnexoTraducao.MoveFirst
'y = rstCarregarAnexoTraducao.RecordCount
    
'    l = 3
'    c = 22
While Not rstCarregarAnexoTraducao.EOF

'    For x = 1 To y
        
        With rstCarregarAnexoTraducao
        
            Cells(intLinha, intColuna).value = .Fields("DESCRICAO")
            Cells(intLinha, intColuna + 1).value = Val(.Fields("VALOR_01"))
            rstCarregarAnexoTraducao.MoveNext
            
        End With
        
        intLinha = intLinha + 1
'    Next x

Wend

CarregarAnexoTraducao_Fim:
    rstCarregarAnexoTraducao.Close
    dbOrcamento.Close
    
    Set dbOrcamento = Nothing
    Set rstCarregarAnexoTraducao = Nothing
    
    Exit Function
CarregarAnexoTraducao_err:
    If Err.Number = "3021" Then
        MsgBox "ATENÇÃO: Não exitem registros de Tradução", vbInformation + vbOKOnly, "Registros de Tradução"
    Else
        MsgBox Err.Number & Space(2) & Err.Description, , "Registros de Tradução"
    End If
    
    Resume CarregarAnexoTraducao_Fim

End Function

Public Function CarregarAnexoRevisao( _
                                    BaseDeDados As String, _
                                    strControle As String, _
                                    strVendedor As String, _
                                    intLinha As Integer, _
                                    intColuna As Integer)
                                    
On Error GoTo CarregarAnexoRevisao_err
                                    
Dim dbOrcamento As DAO.Database
Dim rstCarregarAnexoRevisao As DAO.Recordset

'Dim L As Integer, c As Integer ' L = LINHA | C = COLUNA
'Dim x, y As Integer ' contador de linhas

Set dbOrcamento = DBEngine.OpenDatabase(BaseDeDados, False, False, "MS Access;PWD=" & SenhaBanco)
Set rstCarregarAnexoRevisao = dbOrcamento.OpenRecordset("Select * from OrcamentosAnexos where controle = '" & strControle & "' and Vendedor = '" & strVendedor & "' and PROPRIEDADE = 'REVISÃO'")
    
rstCarregarAnexoRevisao.MoveLast
rstCarregarAnexoRevisao.MoveFirst
'y = rstCarregarAnexoRevisao.RecordCount
    
'    l = 3
'    c = 22

While Not rstCarregarAnexoRevisao.EOF

'    For x = 1 To y
        
        With rstCarregarAnexoRevisao
        
            Cells(intLinha, intColuna).value = .Fields("DESCRICAO")
            Cells(intLinha, intColuna + 1).value = Val(.Fields("VALOR_01"))
            rstCarregarAnexoRevisao.MoveNext
            
        End With
        
        intLinha = intLinha + 1
'    Next x

Wend

CarregarAnexoRevisao_Fim:
    rstCarregarAnexoRevisao.Close
    dbOrcamento.Close
    
    Set dbOrcamento = Nothing
    Set rstCarregarAnexoRevisao = Nothing
    
    Exit Function
CarregarAnexoRevisao_err:
    If Err.Number = "3021" Then
        MsgBox "ATENÇÃO: Não exitem registros de Revisão", vbInformation + vbOKOnly, "Registros de Revisão"
    Else
        MsgBox Err.Number & Space(2) & Err.Description, , "Registros de Revisão"
    End If
    
    Resume CarregarAnexoRevisao_Fim

End Function

Public Function CarregarAnexoDiagramacao( _
                                    BaseDeDados As String, _
                                    strControle As String, _
                                    strVendedor As String, _
                                    intLinha As Integer, _
                                    intColuna As Integer)
                                    
On Error GoTo CarregarAnexoDiagramacao_err
                                    
Dim dbOrcamento As DAO.Database
Dim rstCarregarAnexoDiagramacao As DAO.Recordset

'Dim L As Integer, c As Integer ' L = LINHA | C = COLUNA
'Dim x, y As Integer ' contador de linhas

Set dbOrcamento = DBEngine.OpenDatabase(BaseDeDados, False, False, "MS Access;PWD=" & SenhaBanco)
Set rstCarregarAnexoDiagramacao = dbOrcamento.OpenRecordset("Select * from OrcamentosAnexos where controle = '" & strControle & "' and Vendedor = '" & strVendedor & "' and PROPRIEDADE = 'DIAGRAMAÇÃO'")
    
rstCarregarAnexoDiagramacao.MoveLast
rstCarregarAnexoDiagramacao.MoveFirst
'y = rstCarregarAnexoDiagramacao.RecordCount
    
'    l = 3
'    c = 22
While Not rstCarregarAnexoDiagramacao.EOF

'    For x = 1 To y
        
        With rstCarregarAnexoDiagramacao
        
            Cells(intLinha, intColuna).value = .Fields("DESCRICAO")
            Cells(intLinha, intColuna + 1).value = Val(.Fields("VALOR_01"))
            rstCarregarAnexoDiagramacao.MoveNext
            
        End With
        
        intLinha = intLinha + 1
'    Next x
    
Wend


CarregarAnexoDiagramacao_Fim:
    rstCarregarAnexoDiagramacao.Close
    dbOrcamento.Close
    
    Set dbOrcamento = Nothing
    Set rstCarregarAnexoDiagramacao = Nothing
    
    Exit Function
CarregarAnexoDiagramacao_err:
    If Err.Number = "3021" Then
        MsgBox "ATENÇÃO: Não exitem registros de Diagramação", vbInformation + vbOKOnly, "Registros de Diagramação"
    Else
        MsgBox Err.Number & Space(2) & Err.Description, , "Registros de Diagramação"
    End If
    
    Resume CarregarAnexoDiagramacao_Fim

End Function



Public Function CarregarAnexoLinha( _
                                    BaseDeDados As String, _
                                    strControle As String, _
                                    strVendedor As String, _
                                    intLinha As Integer, _
                                    intColuna As Integer)
                                    
On Error GoTo CarregarAnexoLinha_err
                                    
Dim dbOrcamento As DAO.Database
Dim rstCarregarAnexoLinha As DAO.Recordset

'Dim L As Integer, c As Integer ' L = LINHA | C = COLUNA
'Dim x, y As Integer ' contador de linhas

Set dbOrcamento = DBEngine.OpenDatabase(BaseDeDados, False, False, "MS Access;PWD=" & SenhaBanco)
Set rstCarregarAnexoLinha = dbOrcamento.OpenRecordset("Select * from OrcamentosAnexos where controle = '" & strControle & "' and Vendedor = '" & strVendedor & "' and PROPRIEDADE = 'LINHA'")

rstCarregarAnexoLinha.MoveLast
rstCarregarAnexoLinha.MoveFirst
'y = rstCarregarAnexoLinha.RecordCount
    
'    l = 3
'    c = 12
While Not rstCarregarAnexoLinha.EOF

'    For x = 1 To y
        
        With rstCarregarAnexoLinha
            
            Cells(intLinha, intColuna).value = .Fields("DESCRICAO")
            Cells(intLinha, intColuna + 1).value = .Fields("VALOR_01")
            Cells(intLinha, intColuna + 2).value = .Fields("VALOR_02")
            rstCarregarAnexoLinha.MoveNext
            
        End With
        
        intLinha = intLinha + 1
'    Next x

Wend


CarregarAnexoLinha_Fim:
    rstCarregarAnexoLinha.Close
    dbOrcamento.Close
    
    Set dbOrcamento = Nothing
    Set rstCarregarAnexoLinha = Nothing
    
    Exit Function
CarregarAnexoLinha_err:
    If Err.Number = "3021" Then
        MsgBox "ATENÇÃO: Não exitem registros de Linha de produtos", vbInformation + vbOKOnly, "Registros de Linha de produtos"
    Else
        MsgBox Err.Number & Space(2) & Err.Description, , "Registros de Linha de produtos"
    End If
    
    Resume CarregarAnexoLinha_Fim

End Function


Public Function CarregarAnexoMoeda( _
                                    BaseDeDados As String, _
                                    strControle As String, _
                                    strVendedor As String, _
                                    intLinha As Integer, _
                                    intColuna As Integer)
                                    
On Error GoTo CarregarAnexoMoeda_err
                                    
Dim dbOrcamento As DAO.Database
Dim rstCarregarAnexoMoeda As DAO.Recordset

'Dim L As Integer, c As Integer ' L = LINHA | C = COLUNA
'Dim x, y As Integer ' contador de linhas


Set dbOrcamento = DBEngine.OpenDatabase(BaseDeDados, False, False, "MS Access;PWD=" & SenhaBanco)
Set rstCarregarAnexoMoeda = dbOrcamento.OpenRecordset("Select * from OrcamentosAnexos where controle = '" & strControle & "' and Vendedor = '" & strVendedor & "' and PROPRIEDADE = 'MOEDA'")
    
rstCarregarAnexoMoeda.MoveLast
rstCarregarAnexoMoeda.MoveFirst
'y = rstCarregarAnexoMoeda.RecordCount
    
'    l = 3
'    c = 16

While Not rstCarregarAnexoMoeda.EOF
    
'    For x = 1 To y
        
        With rstCarregarAnexoMoeda
        
            Cells(intLinha, intColuna).value = .Fields("DESCRICAO")
            Cells(intLinha, intColuna + 1).value = .Fields("VALOR_01")
            rstCarregarAnexoMoeda.MoveNext
            
        End With
        
        intLinha = intLinha + 1
'    Next x

Wend


CarregarAnexoMoeda_Fim:
    rstCarregarAnexoMoeda.Close
    dbOrcamento.Close
    
    Set dbOrcamento = Nothing
    Set rstCarregarAnexoMoeda = Nothing
    
    Exit Function
CarregarAnexoMoeda_err:
    
    If Err.Number = "3021" Then
        MsgBox "ATENÇÃO: Não exitem registros de Moeda", vbInformation + vbOKOnly, "Registros de Moeda"
    Else
        MsgBox Err.Number & Space(2) & Err.Description, , "Registros de Moeda"
    End If
    
    Resume CarregarAnexoMoeda_Fim

End Function

Public Function CarregarAnexoVenda( _
                                    BaseDeDados As String, _
                                    strControle As String, _
                                    strVendedor As String, _
                                    intLinha As Integer, _
                                    intColuna As Integer)
                                    
On Error GoTo CarregarAnexoVenda_err
                                    
Dim dbOrcamento As DAO.Database
Dim rstCarregarAnexoVenda As DAO.Recordset

'Dim L As Integer, c As Integer ' L = LINHA | C = COLUNA
'Dim x, y As Integer ' contador de linhas

Set dbOrcamento = DBEngine.OpenDatabase(BaseDeDados, False, False, "MS Access;PWD=" & SenhaBanco)
Set rstCarregarAnexoVenda = dbOrcamento.OpenRecordset("Select * from OrcamentosAnexos where controle = '" & strControle & "' and Vendedor = '" & strVendedor & "' and PROPRIEDADE = 'VENDA'")

rstCarregarAnexoVenda.MoveLast
rstCarregarAnexoVenda.MoveFirst
'y = rstCarregarAnexoVenda.RecordCount
    
While Not rstCarregarAnexoVenda.EOF

'    For x = 1 To y
        
        With rstCarregarAnexoVenda
        
            Cells(intLinha, intColuna).value = .Fields("DESCRICAO")
            Cells(intLinha, intColuna + 1).value = .Fields("VALOR_01")
            rstCarregarAnexoVenda.MoveNext
            
        End With
        
        intLinha = intLinha + 1
'    Next x

Wend

CarregarAnexoVenda_Fim:
    rstCarregarAnexoVenda.Close
    dbOrcamento.Close
    
    Set dbOrcamento = Nothing
    Set rstCarregarAnexoVenda = Nothing
    
    Exit Function
CarregarAnexoVenda_err:
    
    If Err.Number = "3021" Then
        MsgBox "ATENÇÃO: Não exitem registros de Tipo de vendas", vbInformation + vbOKOnly, "Registros de Tipo de vendas"
    Else
        MsgBox Err.Number & Space(2) & Err.Description, , "Registros de Tipo de vendas"
    End If
    
    Resume CarregarAnexoVenda_Fim

End Function


'Public Function CarregarAnexoArquivo( _
'                                    BaseDeDados As String, _
'                                    strControle As String, _
'                                    strVendedor As String, _
'                                    intLinha As Integer, _
'                                    intColuna As Integer)
'
'On Error GoTo CarregarAnexoArquivo_err
'
'Dim dbOrcamento As dao.Database
'Dim rstCarregarAnexoArquivo As dao.Recordset
'
'Dim l As Integer, c As Integer ' L = LINHA | C = COLUNA
'Dim x, y As Integer ' contador de linhas
'
''ARQUIVOS - ( ANEXOS )
'Dim Terminio As Integer
'Dim Inicio As Integer
'
'Set dbOrcamento = DBEngine.OpenDatabase(BaseDeDados, False, False, "MS Access;PWD=" & SenhaBanco)
'Set rstCarregarAnexoArquivo = dbOrcamento.OpenRecordset("Select * from OrcamentosAnexos where controle = '" & strControle & "' and Vendedor = '" & strVendedor & "' and PROPRIEDADE = 'Arquivo'")
'
'If Not rstCarregarAnexoArquivo.EOF Then
'
'    Inicio = CInt(Right(ArquivoInicio, Len(ArquivoInicio) - 1))
'
'    rstCarregarAnexoArquivo.MoveLast
'    rstCarregarAnexoArquivo.MoveFirst
'    Terminio = (rstCarregarAnexoArquivo.RecordCount + Inicio) - 1
'
'    Range(ArquivoControle).Value = rstCarregarAnexoArquivo.RecordCount + Inicio
'
'    For x = Inicio To Terminio
'
'        With rstCarregarAnexoArquivo
'
'            Range(Chr(Asc(Left(ArquivoInicio, 1)) + 1) & x).Select
''            ActiveCell.FormulaR1C1 = vrtSelectedItem
'            Selection.Hyperlinks.Add Range(Chr(Asc(Left(ArquivoInicio, 1)) + 1) & x), "file://" & .Fields("OBS_01")
'            Selection.Font.Size = 12
'
''            Cells(intLinha, intColuna).Value = .Fields("OBS_01")
'            rstCarregarAnexoArquivo.MoveNext
'
'        End With
'
'        intLinha = intLinha + 1
'    Next x
'
'End If
'
'CarregarAnexoArquivo_Fim:
'    rstCarregarAnexoArquivo.Close
'    dbOrcamento.Close
'
'    Set dbOrcamento = Nothing
'    Set rstCarregarAnexoArquivo = Nothing
'
'    Exit Function
'CarregarAnexoArquivo_err:
'    MsgBox Err.Description, , "Anexo Arquivo"
'    Resume CarregarAnexoArquivo_Fim
'
'End Function
'
'
'
