Attribute VB_Name = "basIRibbonControl"
Sub Pesquisar(ByVal control As IRibbonControl)
    frmPesquisar.Show
End Sub

Sub Cadastro(ByVal control As IRibbonControl)
    If Range(GerenteDeContas) <> "" Then
        frmCadastro.Show
    End If
End Sub

Sub AnexosArquivos(ByVal control As IRibbonControl)
    If Range(GerenteDeContas) <> "" Then
        frmAnexosArquivos.Show
    End If
End Sub

Sub EnviarReceber(ByVal control As IRibbonControl)
    frmEnviarReceber.Show
End Sub

Sub Projeto01(ByVal control As IRibbonControl)
Dim strUsuario As String: strUsuario = Range(NomeUsuario)

    If ActiveSheet.Name = strUsuario Then
        Unload frmPojeto
        Exit Sub
    Else
        ProjetoAtual = "C"

        If Range(StatusProjeto) = "Novo" Then
            frmPojeto.Show
        Else
            If Range(ProjetoAtual & "13") <> "" Then
                Exit Sub
            Else
                frmPojeto.Show
            End If
        End If
    End If


End Sub

Sub Projeto02(ByVal control As IRibbonControl)
Dim strUsuario As String: strUsuario = Range(NomeUsuario)

    If ActiveSheet.Name = strUsuario Then
        Unload frmPojeto
        Exit Sub
    Else
        ProjetoAtual = "D"

        If Range(StatusProjeto) = "Novo" Then
            frmPojeto.Show
        Else
            If Range(ProjetoAtual & "13") <> "" Then
                Exit Sub
            Else
                frmPojeto.Show
            End If
        End If
    End If

End Sub

Sub Projeto03(ByVal control As IRibbonControl)
Dim strUsuario As String: strUsuario = Range(NomeUsuario)

    If ActiveSheet.Name = strUsuario Then
        Unload frmPojeto
        Exit Sub
    Else
        ProjetoAtual = "E"

        If Range(StatusProjeto) = "Novo" Then
            frmPojeto.Show
        Else
            If Range(ProjetoAtual & "13") <> "" Then
                Exit Sub
            Else
                frmPojeto.Show
            End If
        End If
    End If

End Sub

Sub Projeto04(ByVal control As IRibbonControl)
Dim strUsuario As String: strUsuario = Range(NomeUsuario)

    If ActiveSheet.Name = strUsuario Then
        Unload frmPojeto
        Exit Sub
    Else
        ProjetoAtual = "F"

        If Range(StatusProjeto) = "Novo" Then
            frmPojeto.Show
        Else
            If Range(ProjetoAtual & "13") <> "" Then
                Exit Sub
            Else
                frmPojeto.Show
            End If
        End If
    End If

End Sub
Sub Projeto05(ByVal control As IRibbonControl)
Dim strUsuario As String: strUsuario = Range(NomeUsuario)

    If ActiveSheet.Name = strUsuario Then
        Unload frmPojeto
        Exit Sub
    Else
        ProjetoAtual = "G"

        If Range(StatusProjeto) = "Novo" Then
            frmPojeto.Show
        Else
            If Range(ProjetoAtual & "13") <> "" Then
                Exit Sub
            Else
                frmPojeto.Show
            End If
        End If
    End If

End Sub

Sub Projeto06(ByVal control As IRibbonControl)
Dim strUsuario As String: strUsuario = Range(NomeUsuario)

    If ActiveSheet.Name = strUsuario Then
        Unload frmPojeto
        Exit Sub
    Else
        ProjetoAtual = "H"

        If Range(StatusProjeto) = "Novo" Then
            frmPojeto.Show
        Else
            If Range(ProjetoAtual & "13") <> "" Then
                Exit Sub
            Else
                frmPojeto.Show
            End If
        End If
    End If

End Sub

Sub Projeto07(ByVal control As IRibbonControl)
Dim strUsuario As String: strUsuario = Range(NomeUsuario)

    If ActiveSheet.Name = strUsuario Then
        Unload frmPojeto
        Exit Sub
    Else
        ProjetoAtual = "I"

        If Range(StatusProjeto) = "Novo" Then
            frmPojeto.Show
        Else
            If Range(ProjetoAtual & "13") <> "" Then
                Exit Sub
            Else
                frmPojeto.Show
            End If
        End If
    End If

End Sub

Sub Projeto08(ByVal control As IRibbonControl)
Dim strUsuario As String: strUsuario = Range(NomeUsuario)

    If ActiveSheet.Name = strUsuario Then
        Unload frmPojeto
        Exit Sub
    Else
        ProjetoAtual = "J"

        If Range(StatusProjeto) = "Novo" Then
            frmPojeto.Show
        Else
            If Range(ProjetoAtual & "13") <> "" Then
                Exit Sub
            Else
                frmPojeto.Show
            End If
        End If
    End If

End Sub

Sub Indices(ByVal control As IRibbonControl)
Dim strBanco As String: strBanco = Range(BancoLocal)
Dim strUsuario As String: strUsuario = Range(NomeUsuario)
Dim strMSG As String
Dim strTitulo As String

If Range(GerenteDeContas) <> "" Then

    If LiberarIndice(strBanco, strUsuario) = False Then
        strMSG = "Ops!!! " & Chr(10) & Chr(13) & Chr(13)
        strMSG = strMSG & "Você não tem permissão para acessar este conteúdo. " & Chr(10) & Chr(13) & Chr(13)
        strTitulo = "Indices de calculos!"

        MsgBox strMSG, vbInformation + vbOKOnly, strTitulo
    Else

        LiberarIndice strBanco, strUsuario
        frmIndices.Show

    End If

End If

End Sub

Sub ENVIAR(ByVal control As IRibbonControl)
    frmEnviar.Show
End Sub

Sub desbloqueio(ByVal control As IRibbonControl)
    DesbloqueioDeGuia SenhaBloqueio
    Sheets("BANCOS").Visible = -1
End Sub

Sub GerarProposta(ByVal control As IRibbonControl)
'    Gerar_Proposta
    MsgBox "EM TESTES", vbInformation + vbOKOnly, "Gerar Proposta."
End Sub

Sub SimuladorCustos(ByVal control As IRibbonControl)
    MsgBox "EM TESTES", vbInformation + vbOKOnly, "Simulador de custos."
End Sub

Sub ControleGrand(ByVal control As IRibbonControl)
    MsgBox "EM TESTES", vbInformation + vbOKOnly, "Cadastro de Grand."
End Sub

Sub EnviarDados(ByVal control As IRibbonControl)
    MsgBox "EM TESTES", vbInformation + vbOKOnly, "Enviar Dados."
End Sub

Sub ReceberDados(ByVal control As IRibbonControl)
    MsgBox "EM TESTES", vbInformation + vbOKOnly, "Receber Dados."
End Sub


Sub modelo_teste(ByVal control As IRibbonControl)
   
Dim strBanco As String: strBanco = Range(BancoLocal)
Dim strControle As String
Dim strUsuario As String

    strControle = InputBox("Informe o numero de controle:", "Numero de controle", "082-14")
    strUsuario = InputBox("Informe o nome do vendedor:", "Nome do vendedor", "azs")

    admLimparAnexos
        
    DesbloqueioDeGuia SenhaBloqueio

    CarregarAnexoLinha strBanco, strControle, strUsuario, 3, 12
    CarregarAnexoMoeda strBanco, strControle, strUsuario, 3, 16
    CarregarAnexoVenda strBanco, strControle, strUsuario, 3, 19
    CarregarAnexoDesconto strBanco, strControle, strUsuario, 3, 22

    CarregarAnexoTraducao strBanco, strControle, strUsuario, 3, 29
    CarregarAnexoRevisao strBanco, strControle, strUsuario, 3, 32
    CarregarAnexoDiagramacao strBanco, strControle, strUsuario, 3, 35

'MarcaTexto InputBox("Informe a seleção:", "seleção", "")
    
End Sub

Sub SelecaoDeArea(ByVal control As IRibbonControl)
'Dim marcacao As String: marcacao = InputBox("Informe a seleção:", "seleção", "")
'
'    DesbloqueioDeGuia SenhaBloqueio
'
'    If marcacao <> "" Then
'        MarcaSelecao marcacao
'    End If
    
End Sub

Sub MenuChoice(control As IRibbonControl)

'DesbloqueioDeGuia SenhaBloqueio
'
'admIntervalosDeEdicaoLimparSelecao Range(BancoLocal)
'
'Select Case control.ID
'
'    Case "menuHistorico"
''        MarcaSelecao ""
'    Case "menuDesconto"
''        MarcaSelecao ""
'    Case "menuReCusto"
'        '' CUSTOS
'        MarcaSelecao "C37:J57"
'    Case "menuCancelado"
''        MarcaSelecao ""
'    Case "menuExcluido"
''        MarcaSelecao ""
'    Case "menuVendido"
''        MarcaSelecao ""
'    Case "menuNovo"
'
'        '' ORÇAMENTO
'        MarcaSelecao "C4,C5,G3:G4,C6,C8:J10,C13:J15,C17:J23,C61:J61,C23:J23"
'
'        '' ROYALTY PERCENTUAL
'        MarcaSelecao "C21:J21"
'
'        '' ROYALTY ESPECIE
'        MarcaSelecao "C22:J22"
'
'        '' IMPRESSÃO
'        MarcaSelecao "C25:J29,B31:J34"
'
'        '' DESCONTOS
'        MarcaSelecao "C61:J61"
'
'        '' PREÇO MKT
'        MarcaSelecao "C73:J73"
'
'        '' PREÇO COMPRA
'        MarcaSelecao "C80:J80"
'
'        '' DESCONTO COMPRA
'        MarcaSelecao "C79:J79"
'
'    Case "menuCusto"
'
'        '' CUSTOS
'        MarcaSelecao "C37:J57"
'
'    Case "menuOrcamento"
'
'        '' CUSTOS
'        MarcaSelecao "C37:J57"
'
'    Case "menuPreco"
'
'        '' ORÇAMENTO
'        MarcaSelecao "C4,C5,G3:G4,C6,C8:J10,C13:J15,C17:J23,C61:J61,C23:J23"
'
'        '' ROYALTY PERCENTUAL
'        MarcaSelecao "C21:J21"
'
'        '' ROYALTY ESPECIE
'        MarcaSelecao "C22:J22"
'
'        '' IMPRESSÃO
'        MarcaSelecao "C25:J29,B31:J34"
'
'        '' DESCONTOS
'        MarcaSelecao "C61:J61"
'
'        '' PREÇO MKT
'        MarcaSelecao "C73:J73"
'
'        '' PREÇO COMPRA
'        MarcaSelecao "C80:J80"
'
'        '' DESCONTO COMPRA
'        MarcaSelecao "C79:J79"
'
'End Select

End Sub

Sub Administracao(ByVal control As IRibbonControl)
    frmAdministracao.Show
    Sheets("BANCOS").Visible = 2
End Sub

Sub teste_formatos()

Dim ws As Worksheet
Set ws = Worksheets(ActiveSheet.Name)

For Each cLoc In ws.Range("Projetos")
    If cLoc <> "" Then
        MsgBox cLoc
    End If
Next cLoc

End Sub
