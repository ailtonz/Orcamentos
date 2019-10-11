Attribute VB_Name = "basExcel"
Option Explicit

Public Function BloqueioDeGuia(strSenha As String)
    ActiveSheet.Protect strSenha
End Function

Public Function DesbloqueioDeGuia(strSenha As String)
    ActiveSheet.Unprotect strSenha
End Function

Function LimparTemplate(Selecao As String, Conteudo As Variant)
    Range(Selecao).Select
    Selection.value = Conteudo
End Function

Public Function OcultarLinhas(LinhaInicio As Integer, LinhaFinal As Integer, ocultar As Boolean)
    Rows(CStr(LinhaInicio) & ":" & CStr(LinhaFinal)).Select
    Selection.EntireRow.Hidden = ocultar
End Function

Public Function IntervaloEditacaoCriar(Titulo As String, Selecao As String, Optional MarcarSelecao As Boolean)
On Error GoTo IntervaloEditacaoCriar_err
'MarcarSelecao = False
'On Error Resume Next
    
    If Not IntervaloEditacaoExiste(Titulo) Then
       
        ActiveSheet.Protection.AllowEditRanges.Add Title:=Titulo, Range:=Range(Selecao)
        
        If Not MarcarSelecao Then
            DesmarcaTexto Selecao
        Else
            MarcaTexto Selecao
        End If
    
    End If
    
IntervaloEditacaoCriar_Fim:
Exit Function
    
IntervaloEditacaoCriar_err:
    MsgBox Err.Description
    Resume IntervaloEditacaoCriar_Fim
End Function

Public Function IntervaloEditacaoRemover(IntervaloDeEdicao As String, Optional MarcarSelecao As String)
    Dim AER As AllowEditRange
    
    For Each AER In ActiveSheet.Protection.AllowEditRanges
        If AER.Title = IntervaloDeEdicao Then
            AER.Delete
            If MarcarSelecao <> "" Then
                MarcaTexto MarcarSelecao
            End If
            
        End If
    Next AER

End Function

Public Function IntervaloEditacaoRemoverTodos()
Dim AER As AllowEditRange
Dim x As Integer

    x = ActiveSheet.Protection.AllowEditRanges.Count
        
    For Each AER In ActiveSheet.Protection.AllowEditRanges
        If x > 0 Then
            AER.Delete
        End If
    Next AER
        
    Set AER = Nothing

End Function

Public Function IntervaloEditacaoExiste(strTitulo As String) As Boolean: IntervaloEditacaoExiste = False
Dim AER As AllowEditRange
Dim x As Integer

    x = ActiveSheet.Protection.AllowEditRanges.Count
        
    For Each AER In ActiveSheet.Protection.AllowEditRanges
        If AER.Title = strTitulo Then
            IntervaloEditacaoExiste = True
        End If
    Next AER
        
    Set AER = Nothing

End Function

Public Function MarcaSelecao(Selecao As String)
''####################
''     BRANCO
'' LINK: http://msdn.microsoft.com/en-us/library/cc296089(v=office.12).aspx
''####################
    
    Range(Selecao).Select
    
    With Selection.Interior
'        .Pattern = xlSolid
'        .PatternColorIndex = xlAutomatic
'        .ColorIndex = 36
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
    
End Function

Public Function MarcaTexto(Selecao As String)
    
    Range(Selecao).Select
    
    With Selection.Interior
'''####################
'''     BRANCO
''' LINK: http://msdn.microsoft.com/en-us/library/cc296089(v=office.12).aspx
'''####################
'
'        .Pattern = xlSolid
'        .PatternColorIndex = xlAutomatic
'        .ColorIndex = 36
'        .TintAndShade = 0
'        .PatternTintAndShade = 0
        
'''####################
'''     SIMPLES
'''####################

        .Pattern = xlGray8
        .PatternColorIndex = xlAutomatic
        .ColorIndex = xlAutomatic
'        .TintAndShade = 0
'        .PatternTintAndShade = 0
        
    End With
    
End Function

Public Function DesmarcaTexto(Selecao As String)
    
'Application.Wait DateAdd("s", 10, Now)
    
    Range(Selecao).Select
    
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
'        .ThemeColor = xlThemeColorDark1
'        .TintAndShade = 0
'        .PatternTintAndShade = 0
    End With
    
End Function

Public Sub InserirConteudo(Linha As Long, Coluna As Long, Conteudo As String)
    Cells(Linha, Coluna).Select
    Cells(Linha, Coluna).value = Conteudo
End Sub

Public Function SelecionarGuiaAtual()
    Sheets(ActiveSheet.Name).Select
End Function

Public Function PesquisaNomeGuia(sGuia As String) As Boolean
Dim s As Integer
    For s = 1 To Sheets.Count
        If Sheets(s).Name = sGuia Then
            PesquisaNomeGuia = True
        End If
    Next
End Function

Public Function ContarAreaPreechida(area As Range) As Long
    Dim celula As Range, contador As Long
    contador = 0
    For Each celula In area
        If celula <> "" Then
            contador = contador + 1
        End If
    Next
    ContarAreaPreechida = contador
End Function

Function MarcarObrigatorio(ByVal strCelula As String, Marcar As Boolean)
'' Marcar celula obrigatoria quando estiver vasia
    If Range(strCelula) = "" Or Range(strCelula) = 0 Then
        Range(strCelula).Select
        If Marcar Then
            With Selection.Interior
                .Pattern = xlSolid
                .PatternColorIndex = xlAutomatic
                .Color = 255
                .TintAndShade = 0
                .PatternTintAndShade = 0
            End With
        Else
            With Selection.Interior
                .Pattern = xlNone
                .TintAndShade = 0
                .PatternTintAndShade = 0
            End With
        End If
    End If
End Function

Sub MoverCursor(posicao As String)

Select Case posicao
    
    Case "cima"
        ActiveCell.Offset(-1, 0).Select

    Case "baixo"
        ActiveCell.Offset(1, 0).Select
    
    Case "direita"
        ActiveCell.Offset(0, 1).Select

    Case "esquerda"
        ActiveCell.Offset(0, -1).Select
    
End Select

End Sub



'Public Function MarcarObrigatorio(Selecao As String)
'
'    Range(Selecao).Select
'    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
'    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
'    With Selection.Borders(xlEdgeLeft)
'        .LineStyle = xlContinuous
'        .Color = -16776961
'        .TintAndShade = 0
'        .Weight = xlThick
'    End With
'    With Selection.Borders(xlEdgeTop)
'        .LineStyle = xlContinuous
'        .Color = -16776961
'        .TintAndShade = 0
'        .Weight = xlThick
'    End With
'    With Selection.Borders(xlEdgeBottom)
'        .LineStyle = xlContinuous
'        .Color = -16776961
'        .TintAndShade = 0
'        .Weight = xlThick
'    End With
'    With Selection.Borders(xlEdgeRight)
'        .LineStyle = xlContinuous
'        .Color = -16776961
'        .TintAndShade = 0
'        .Weight = xlThick
'    End With
'    With Selection.Borders(xlInsideVertical)
'        .LineStyle = xlContinuous
'        .ColorIndex = 0
'        .TintAndShade = 0
'        .Weight = xlThin
'    End With
'    Selection.Borders(xlInsideHorizontal).LineStyle = xlNone
'
'End Function
