Attribute VB_Name = "basEmail"
Option Explicit

Function EnviarOrcamentos(strEmail As String, strAssunto As String, strArquivo As String, strConteudo As String)
On Error GoTo EnviarOrcamentos_err
' Works in Excel 2000, Excel 2002, Excel 2003, Excel 2007, Excel 2010, Outlook 2000, Outlook 2002, Outlook 2003, Outlook 2007, Outlook 2010.
' This example sends the last saved version of the Activeworkbook object .
    
    Dim OutApp As Object
    Dim OutMail As Object
    
    Dim L As Integer, c As Integer ' L = LINHA | C = COLUNA
    Dim x As Integer ' contador de linhas
    

    Set OutApp = CreateObject("Outlook.Application")
    Set OutMail = OutApp.CreateItem(0)

    On Error Resume Next
   ' Change the mail address and subject in the macro before you run it.
    With OutMail
'        .From = "tania.silva@springer.com"
        .To = strEmail
        .cc = ""
        .BCC = ""
        .Subject = strAssunto
        .Body = strConteudo
        .Attachments.Add strArquivo
        .Send
    End With
    On Error GoTo 0
    

EnviarOrcamentos_Fim:
    Set OutMail = Nothing
    Set OutApp = Nothing
  
    Exit Function
EnviarOrcamentos_err:
    MsgBox Err.Description
    Resume EnviarOrcamentos_Fim
    
End Function

Function EnviarEmail(strEmail As String, _
                     strAssunto As String, _
                     Anexo As Boolean, _
                     Optional BaseDeDados As String, _
                     Optional strConsulta As String, _
                     Optional strCampo As String)

On Error GoTo EnviarEmail_err
        
    Dim OutApp As Object
    Set OutApp = CreateObject("Outlook.Application")
    
    Dim OutMail As Object
    Set OutMail = OutApp.CreateItem(0)

    On Error Resume Next
   ' Change the mail address and subject in the macro before you run it.
    With OutMail
'        .From = "tania.silva@springer.com"
        .To = strEmail
        .cc = ""
        .BCC = ""
        .Subject = strAssunto
'        .Body = "Hello World!"
'        .Attachments.Add ActiveWorkbook.FullName
        ' You can add other files by uncommenting the following line.
        
        If Anexo = True Then
        
            '   BASE DE DADOS
            Dim dbOrcamento As DAO.Database
            Set dbOrcamento = DBEngine.OpenDatabase(BaseDeDados, False, False, "MS Access;PWD=" & SenhaBanco)
            
            '   CONSULTA
            Dim rstOrcamentosAnexos As DAO.Recordset
            Set rstOrcamentosAnexos = dbOrcamento.OpenRecordset(strConsulta)
            
            While Not rstOrcamentosAnexos.EOF
                .Attachments.Add rstOrcamentosAnexos.Fields(strCampo).value
                rstOrcamentosAnexos.MoveNext
            Wend
            
        End If

        ' In place of the following statement, you can use ".Display" to
        ' display the mail.
        
        .Send
    End With
    On Error GoTo 0
    

EnviarEmail_Fim:
    Set OutMail = Nothing
    Set OutApp = Nothing
    
    If Anexo = True Then
        Set dbOrcamento = Nothing
        Set rstOrcamentosAnexos = Nothing
    End If
  
    Exit Function
EnviarEmail_err:
    MsgBox Err.Description
    Resume EnviarEmail_Fim
    
End Function
