Attribute VB_Name = "basUserForm"
Option Explicit

Public Function UserFormDesbloqueioDeFuncoes(BaseDeDados As String, frm As UserForm, strSQL As String, strCampo As String)
On Error GoTo UserFormDesbloqueioDeFuncoes_err

Dim dbOrcamento         As DAO.Database
Dim rstUserFormDesbloqueioDeFuncoes   As DAO.Recordset
Dim RetVal              As Variant
Dim Ctrl                As control

RetVal = Dir(BaseDeDados)

If RetVal = "" Then

    UserFormDesbloqueioDeFuncoes = False
    
Else
        
    Set dbOrcamento = DBEngine.OpenDatabase(BaseDeDados, False, False, "MS Access;PWD=" & SenhaBanco)
    Set rstUserFormDesbloqueioDeFuncoes = dbOrcamento.OpenRecordset(strSQL)
        
    While Not rstUserFormDesbloqueioDeFuncoes.EOF
        For Each Ctrl In frm.Controls
            If TypeName(Ctrl) = "CommandButton" Then
                If Right(Ctrl.Name, Len(Ctrl.Name) - 3) = rstUserFormDesbloqueioDeFuncoes.Fields(strCampo) Then
                    Ctrl.Enabled = True
                End If
            ElseIf TypeName(Ctrl) = "ListBox" Then
                If Right(Ctrl.Name, Len(Ctrl.Name) - 3) = rstUserFormDesbloqueioDeFuncoes.Fields(strCampo) Then
                    Ctrl.Enabled = True
                End If
            End If
        Next
        rstUserFormDesbloqueioDeFuncoes.MoveNext
    Wend
    
    rstUserFormDesbloqueioDeFuncoes.Close
    dbOrcamento.Close
    
    Set dbOrcamento = Nothing
    Set rstUserFormDesbloqueioDeFuncoes = Nothing
    
End If

UserFormDesbloqueioDeFuncoes_Fim:
  
    Exit Function
UserFormDesbloqueioDeFuncoes_err:
    UserFormDesbloqueioDeFuncoes = False
    MsgBox Err.Description
    Resume UserFormDesbloqueioDeFuncoes_Fim
End Function


Sub AtualizarProcesso(Percentual As Single, frm As UserForm) 'variável reservada para ser %

    With frm 'With usa o frmprocesso para as ações abaixo
    'sem ter que repetir o nome do objeto frmprocesso

        ' Atualiza o Título do Quadro que comporta a barra para %
'        .FrameProcesso.Caption = Format(Percentual, "0%")

        ' Atualza o tamanho da Barra (label)
        .lblProcesso.Width = Percentual * (100 - 10)
    End With 'final do uso de frmprocesso diretamente
    
    'Habilita o userform para ser atualizado
    DoEvents
End Sub

Function ListBoxChecarSelecao(frm As UserForm, strListBoxNome As String) As Boolean: ListBoxChecarSelecao = False
Dim Ctrl As control
Dim intCurrentRow As Integer

For Each Ctrl In frm.Controls
    If TypeName(Ctrl) = "ListBox" Then
        If Ctrl.Name = strListBoxNome Then
        
            For intCurrentRow = 0 To Ctrl.ListCount - 1
                If Ctrl.Selected(intCurrentRow) = True Then
                    ListBoxChecarSelecao = True
                    Exit Function
                End If
            Next intCurrentRow
            
        End If
    End If
Next

End Function

Function ListBoxUpdate(wsGuia As String, strListagem As String, frm As UserForm, NomeLista As String)
Dim cLoc As Range
Dim ws As Worksheet
Set ws = Worksheets(wsGuia)

Dim Ctrl As control
Dim x As Long: x = 3
Dim y As Long: y = 1

For Each Ctrl In frm.Controls
    If TypeName(Ctrl) = "ListBox" Then
        If Ctrl.Name = NomeLista Then
            Ctrl.Clear
            For Each cLoc In ws.Range(strListagem)
              Ctrl.AddItem cLoc.value & " | " & cLoc.Cells(x, y)
              y = y + 1
            Next cLoc
        End If
    End If
Next

Set ws = Nothing

End Function

Public Function ListBoxCarregar(BaseDeDados As String, frm As UserForm, NomeLista As String, strCampo As String, strSQL As String)
On Error GoTo ListBoxCarregar_err

Dim dbOrcamento         As DAO.Database
Dim rstListBoxCarregar   As DAO.Recordset
Dim RetVal              As Variant

Dim Ctrl                As control

RetVal = Dir(BaseDeDados)

If RetVal = "" Then

    ListBoxCarregar = False
    
Else
       
    Set dbOrcamento = DBEngine.OpenDatabase(BaseDeDados, False, False, "MS Access;PWD=" & SenhaBanco)
    Set rstListBoxCarregar = dbOrcamento.OpenRecordset(strSQL)
    
    For Each Ctrl In frm.Controls
        If TypeName(Ctrl) = "ListBox" Then
            If Ctrl.Name = NomeLista Then
                Ctrl.Clear
                While Not rstListBoxCarregar.EOF
                    Ctrl.AddItem rstListBoxCarregar.Fields(strCampo)
                    rstListBoxCarregar.MoveNext
                Wend
            End If
        End If
    Next
    
    rstListBoxCarregar.Close
    dbOrcamento.Close
    
    Set dbOrcamento = Nothing
    Set rstListBoxCarregar = Nothing
    
End If

ListBoxCarregar_Fim:
  
    Exit Function
ListBoxCarregar_err:
    ListBoxCarregar = False
    MsgBox Err.Description
    Resume ListBoxCarregar_Fim
End Function




Public Function ListBoxCarregarADO(strLocal As infBanco, frm As UserForm, NomeLista As String, strCampo As String, strSQL As String)
On Error GoTo ListBoxCarregar_err

Dim connection As New ADODB.connection

Dim rstListBoxCarregar As ADODB.Recordset
Set rstListBoxCarregar = New ADODB.Recordset

Dim Ctrl                As control

''Is Internet Connected
If IsInternetConnected() = True Then
    Set connection = OpenConnection(strLocal)
    '' Is Connected
    If connection.State = 1 Then
        
        Call rstListBoxCarregar.Open(strSQL, connection, adOpenStatic, adLockOptimistic)
            
        For Each Ctrl In frm.Controls
        If TypeName(Ctrl) = "ListBox" Then
            If Ctrl.Name = NomeLista Then
                Ctrl.Clear
                While Not rstListBoxCarregar.EOF
                    Ctrl.AddItem rstListBoxCarregar.Fields(strCampo)
                    rstListBoxCarregar.MoveNext
                Wend
            End If
        End If
        Next
        
    Else
        MsgBox "Falha na conexão com o banco de dados!", vbCritical + vbOKOnly, "Falha na conexão com o banco."
    End If
    connection.Close
Else
    ' no connected
    MsgBox "SEM INTERNET.", vbOKOnly + vbExclamation
End If

ListBoxCarregar_Fim:
  
    Exit Function
ListBoxCarregar_err:
    MsgBox Err.Description
    Resume ListBoxCarregar_Fim
End Function


Function ComboBoxUpdate(wsGuia As String, lstListagem As String, cbo As ComboBox)
Dim cLoc As Range
Dim ws As Worksheet
Set ws = Worksheets(wsGuia)

cbo.Clear

For Each cLoc In ws.Range(lstListagem)
  With cbo
    .AddItem cLoc.value
  End With
Next cLoc

End Function

Public Function ComboBoxCarregar(BaseDeDados As String, cbo As ComboBox, strCampo As String, strSQL As String) As Boolean: ComboBoxCarregar = True
On Error GoTo ComboBoxCarregar_err
Dim dbOrcamento As DAO.Database
Dim rstComboBoxCarregar As DAO.Recordset
Dim RetVal As Variant

RetVal = Dir(BaseDeDados)

If RetVal = "" Then

    ComboBoxCarregar = False
    
Else
    
    Set dbOrcamento = DBEngine.OpenDatabase(BaseDeDados, False, False, "MS Access;PWD=" & SenhaBanco)
    Set rstComboBoxCarregar = dbOrcamento.OpenRecordset(strSQL)
    
    cbo.Clear
    
    While Not rstComboBoxCarregar.EOF
        cbo.AddItem rstComboBoxCarregar.Fields(strCampo)
        rstComboBoxCarregar.MoveNext
    Wend
        
    rstComboBoxCarregar.Close
    dbOrcamento.Close
    
    Set dbOrcamento = Nothing
    Set rstComboBoxCarregar = Nothing
    
End If

ComboBoxCarregar_Fim:
  
    Exit Function
ComboBoxCarregar_err:
    ComboBoxCarregar = False
    MsgBox Err.Description
    Resume ComboBoxCarregar_Fim
End Function

