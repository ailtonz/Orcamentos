Attribute VB_Name = "mod_testes"
Option Explicit

Sub UpdateSystem(sUserName As String)

loadBancos
listarAtualizacoes Banco(0), sUserName

 End Sub

'Sub listUpdate()
'Dim myArray() As String
'
'Dim numLimit As Long: numLimit = 20
'
'
'ReDim myArray(1 To numLimit)
'
'For x = 1 To numLimit
'    myArray(x) = x
'Next x
'
'
'MsgBox UBound(myArray), vbInformation, "LIMITE"
'MsgBox LBound(myArray), vbInformation, "INICIO"
'
'End Sub

Sub listUpdate(myArray() As String)
Dim x As Integer

For x = LBound(myArray) To UBound(myArray)
    MsgBox myArray(x)
Next x

End Sub

Sub teste_UpdateSystem(sDescricao As String, sScript As String)
Dim sDataAtualizacao As String: sDataAtualizacao = Controle

loadBancos
admCadastroAtualizacao Banco(0), sDataAtualizacao, 0, ""
admCadastroAtualizacaoScript Banco(0), sDataAtualizacao, sDescricao, sScript

End Sub

Sub teste_listarAtualizacoes()

'loadBancos
'listarAtualizacoes Banco(0), "PATRICIA MIRRA"

End Sub

Function listarAtualizacoes(sBanco As infBanco, sUserName As String)
On Error GoTo listarAtualizacoes_err
Dim cnnServidor As New ADODB.connection
Dim rst As New ADODB.Recordset
Dim myArray() As String
Dim x As Integer
Dim y As Integer
Dim strSQL As String
strSQL = "SELECT * FROM qryUpdateSystem WHERE UserNames like '%" & sUserName & "%' ORDER BY id "

    Set cnnServidor = OpenConnection(sBanco)
    If cnnServidor.State = 1 Then
        Call rst.Open(strSQL, cnnServidor, adOpenStatic, adLockOptimistic)
        
        Do While Not rst.EOF
           
            admUpdateSystem sBanco, Banco(1), rst.Fields("id")
            
            admUpdateSystemRemoveUser sBanco, sUserName, rst.Fields("id")
            
            rst.MoveNext
        Loop

    Else
        MsgBox "Falha na conexão com o banco de dados!", vbCritical + vbOKOnly, "ERROR DE FUNÇÃO: listarAtualizacoes"
    End If

    cnnServidor.Close
    
listarAtualizacoes_Fim:
    Set cnnServidor = Nothing
    Set rst = Nothing

    Exit Function
listarAtualizacoes_err:
    MsgBox Err.Description
    Resume listarAtualizacoes_Fim

End Function

Function admUpdateSystem(strServidor As infBanco, strLocal As infBanco, idAtualizacao As Integer)
On Error GoTo admUpdateSystem_err
Dim cnnServidor As New ADODB.connection
Dim cnnLocal As New ADODB.connection
Dim rst As New ADODB.Recordset
Dim rstLocal As New ADODB.Recordset
Dim cmdLocal As New ADODB.Command
Dim strSQL As String

strSQL = "SELECT * FROM qryUpdateScripts WHERE codRelacao = " & idAtualizacao & ""

    Set cnnServidor = OpenConnection(strServidor)
    If cnnServidor.State = 1 Then
        Call rst.Open(strSQL, cnnServidor, adOpenStatic, adLockOptimistic)
        
        Set cnnLocal = OpenConnection(strLocal)
        If cnnLocal.State = 1 Then
        
            Do While Not rst.EOF
                cmdLocal.ActiveConnection = cnnLocal
                cmdLocal.CommandType = adCmdText
                cmdLocal.CommandText = rst.Fields("SCRIPT").value
                Set rstLocal = cmdLocal.Execute
                rst.MoveNext
            Loop
        
        '            Else
'                MsgBox "Falha na conexão com o banco de dados!", vbCritical + vbOKOnly, "ERROR DE FUNÇÃO: admUpdateSystem"
        End If
        
'    Else
'        MsgBox "Falha na conexão com o banco de dados!", vbCritical + vbOKOnly, "ERROR DE FUNÇÃO: admUpdateSystem"
    End If
    
    cnnServidor.Close
    cnnLocal.Close
    
admUpdateSystem_Fim:
    Set cnnServidor = Nothing
    Set cnnLocal = Nothing
    Set rst = Nothing
    Set rstLocal = Nothing
    Set cmdLocal = Nothing
    
    Exit Function
admUpdateSystem_err:
    MsgBox Err.Description
    Resume admUpdateSystem_Fim
    
End Function

Function admCadastroAtualizacao(strBanco As infBanco, sAtualizacao As String, sIDSubCategoria As String, sObs As String) As Boolean: admCadastroAtualizacao = True
On Error GoTo admCadastroAtualizacao_err
Dim cnn As New ADODB.connection
Set cnn = OpenConnection(strBanco)
Dim rst As ADODB.Recordset
Dim cmd As ADODB.Command

Set cmd = New ADODB.Command
With cmd
    .ActiveConnection = cnn
    .CommandText = "admCategoriaNovo"
    .CommandType = adCmdStoredProc
    .Parameters.Append .CreateParameter("@NM_CATEGORIA", adVarChar, adParamInput, 100, sAtualizacao)
    .Parameters.Append .CreateParameter("@ID_SUBCATEGORIA", adVarChar, adParamInput, 10, sIDSubCategoria)
    .Parameters.Append .CreateParameter("@NM_OBSERVACOES", adVarChar, adParamInput, 2000, sObs)
        
    Set rst = .Execute
    
End With
cnn.Close

admCadastroAtualizacao_Fim:
    Set cnn = Nothing
    Set rst = Nothing
    Set cmd = Nothing
    
    Exit Function
admCadastroAtualizacao_err:
    admCadastroAtualizacao = False
    MsgBox Err.Description
    Resume admCadastroAtualizacao_Fim

End Function

Function admCadastroAtualizacaoScript(strBanco As infBanco, sAtualizacao As String, sDescricao As String, sScript As String) As Boolean: admCadastroAtualizacaoScript = True
On Error GoTo admCadastroAtualizacaoScript_err
Dim cnn As New ADODB.connection
Set cnn = OpenConnection(strBanco)
Dim rst As ADODB.Recordset
Dim cmd As ADODB.Command

Set cmd = New ADODB.Command
With cmd
    .ActiveConnection = cnn
    .CommandText = "admCategoriaScript"
    .CommandType = adCmdStoredProc
    .Parameters.Append .CreateParameter("@ID_SUBCATEGORIA", adVarChar, adParamInput, 10, sAtualizacao)
    .Parameters.Append .CreateParameter("@NM_SUBCATEGORIA", adVarChar, adParamInput, 100, sDescricao)
    .Parameters.Append .CreateParameter("@NM_SELECAO", adVarChar, adParamInput, 2000, sScript)
        
    Set rst = .Execute
    
End With
cnn.Close

admCadastroAtualizacaoScript_Fim:
    Set cnn = Nothing
    Set rst = Nothing
    Set cmd = Nothing
    
    Exit Function
admCadastroAtualizacaoScript_err:
    admCadastroAtualizacaoScript = False
    MsgBox Err.Description
    Resume admCadastroAtualizacaoScript_Fim

End Function

Function ListarUsuariosAtivos(strBanco As infBanco) As String
Dim connection As New ADODB.connection
Dim rst As New ADODB.Recordset
Dim sListagem As String

    Set connection = OpenConnection(strBanco)
    If connection.State = 1 Then
        Call rst.Open("Select usuario from qryUsuarios WHERE (((qryUsuarios.ExclusaoVirtual)=0))", connection, adOpenStatic, adLockOptimistic)
        If Not rst.EOF Then
           Do While Not rst.EOF
                sListagem = sListagem & "|" & rst.Fields("usuario").value
                rst.MoveNext
            Loop
            ListarUsuariosAtivos = Left(Right(sListagem, Len(sListagem) - 1), Len(sListagem) - 1)
            
        End If
    Else
        MsgBox "Falha na conexão com o banco de dados!", vbCritical + vbOKOnly, "Falha na conexão com o banco."
    End If
    connection.Close
End Function

Function admUpdateSystemRemoveUser(strBanco As infBanco, sUser As String, sID As String) As Boolean: admUpdateSystemRemoveUser = True
On Error GoTo admUpdateSystemRemoveUser_err
Dim cnn As New ADODB.connection
Set cnn = OpenConnection(strBanco)
Dim rst As ADODB.Recordset
Dim cmd As ADODB.Command

Set cmd = New ADODB.Command
With cmd
    .ActiveConnection = cnn
    .CommandText = "admUpdateSystemRemoveUser"
    .CommandType = adCmdStoredProc
    .Parameters.Append .CreateParameter("@NM_USER", adVarChar, adParamInput, 50, sUser)
    .Parameters.Append .CreateParameter("@ID", adVarChar, adParamInput, 10, sID)
        
    Set rst = .Execute
    
End With
cnn.Close

admUpdateSystemRemoveUser_Fim:
    Set cnn = Nothing
    Set rst = Nothing
    Set cmd = Nothing
    
    Exit Function
admUpdateSystemRemoveUser_err:
    admUpdateSystemRemoveUser = False
    MsgBox Err.Description
    Resume admUpdateSystemRemoveUser_Fim

End Function
