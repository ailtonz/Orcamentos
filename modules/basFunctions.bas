Attribute VB_Name = "basFunctions"
Option Explicit

Function RoundUp(ByVal value As Double) As Long
   If value > 0 Then value = value * -1
   value = Int(value)
   value = Abs(value)
   RoundUp = value
End Function

Sub ListProcedures()
Dim objCode As VBIDE.CodeModule
Dim objComponent As VBIDE.VBComponent

Dim iLine As Integer
Dim sProcName As String
Dim pk As vbext_ProcKind

' Iterate through each component in the project.
For Each objComponent In ActiveWorkbook.VBProject.VBComponents
    
    ' Find the code module for the project.
    Set objCode = objComponent.CodeModule
    
    Select Case objComponent.Type
    
        Case vbext_ct_ClassModule, vbext_ct_Document
        
        Case vbext_ct_MSForm
        
        Case vbext_ct_StdModule
        
            iLine = 1
            Do While iLine < objCode.CountOfLines
                sProcName = objCode.ProcOfLine(iLine, pk)
                If sProcName <> "" Then
                    Saida objComponent.Name & ": " & sProcName, Controle & ".txt"
    '                Debug.Print objComponent.Name & ": " & sProcName
                    iLine = iLine + objCode.ProcCountLines(sProcName, pk)
                Else
                    ' This line has no procedure, so go to the next line.
                    iLine = iLine + 1
                End If
            Loop
        
    End Select
        
        Set objCode = Nothing
        Set objComponent = Nothing

'    Debug.Print objComponent.Name

Next objComponent

'MsgBox ActiveWorkbook.VBProject.VBComponents("basZip").CodeModule

End Sub

Public Function SelecionarBanco() As String
Dim fd As Office.FileDialog
Dim strArq As String
    
    On Error GoTo SelecionarBanco_err
    
    'Diálogo de selecionar arquivo - Office
    Set fd = Application.FileDialog(msoFileDialogFilePicker)
    fd.Filters.Clear
    fd.Filters.Add "BDs do Access", "*.MDB;*.MDE"
    fd.Title = "Localize a fonte de dados"
    fd.AllowMultiSelect = False
    If fd.Show = -1 Then
        strArq = fd.SelectedItems(1)
    End If
        
    If strArq <> "" Then SelecionarBanco = strArq

SelecionarBanco_Fim:
    Exit Function

SelecionarBanco_err:
    MsgBox Err.Description
    Resume SelecionarBanco_Fim

End Function

Public Function Controle() As String
    Controle = Right(Year(Now()), 2) & Format(Month(Now()), "00") & Format(Day(Now()), "00") & "-" & Format(hour(Now()), "00") & Format(Minute(Now()), "00")
End Function

Public Function DivisorDeTexto(Texto As String, divisor As String, Indice As Integer) As String
On Error Resume Next
Dim Matriz As Variant
    
    Matriz = Array()
    Matriz = Split(Texto, divisor)
    DivisorDeTexto = Trim(CStr(Matriz(Indice)))

End Function

Public Function Saida(strConteudo As String, strArquivo As String)
    Open CreateObject("WScript.Shell").SpecialFolders("Desktop") & "\" & strArquivo For Append As #1
    Print #1, strConteudo
    Close #1
End Function

Function ListarDiretorio(strCaminho As String, strExtensao As String) As String
Dim resultado As Variant
Dim Arquivos As Variant
Dim TamVarNome As Variant

Arquivos = Dir(strCaminho & "\" & strExtensao, vbArchive) ' Recupera a primeira  entrada.
    
''' CHECA A EXISTENCIA DE ARQUIVOS.
If Len(Arquivos) > 0 Then
    Do While Arquivos <> "" ' Inicia o loop.
        resultado = Arquivos & ";" & resultado
        Arquivos = Dir ' Obtém a próxima entrada.
    Loop
    TamVarNome = Mid(resultado, 1, Val(Len(resultado)) - 1)
    ListarDiretorio = TamVarNome
End If

End Function

Function getLineTextFile(myFileName As String, myLine As String) As String
    Dim FileNum As Long
     
    FileNum = FreeFile
    Close FileNum
    Open myFileName For Input As FileNum
    Do While Not EOF(FileNum)
        Line Input #FileNum, myLine
        getLineTextFile = myLine
    Loop
    
    Close FileNum
    
End Function

Function fileExist(filePath As String) As Boolean
    fileExist = IIf(Dir(filePath) <> vbNullString, True, False)
End Function

Function getFileStatus(filespec) As Boolean
   Dim fso, MSG
   Set fso = CreateObject("Scripting.FileSystemObject")
   If (fso.FileExists(filespec)) Then
      getFileStatus = True
   Else
      getFileStatus = False
   End If

End Function
Function getFileStep(strNomeArquivo As String) As String
''' Função particular ao projeto "Orçamentos" Responsavel por informar qual o passo do arquivo.
    getFileStep = Right(getFileName(pathWorkSheetAddress & strNomeArquivo), Len(getFileName(pathWorkSheetAddress & strNomeArquivo)) - 14)
End Function

Public Function getPath(sPathIn As String) As String
'''Esta função irá retornar apenas o path de uma string que contenha o path e o nome do arquivo:
Dim i As Integer

  For i = Len(sPathIn) To 1 Step -1
     If InStr(":\", Mid$(sPathIn, i, 1)) Then Exit For
  Next

  getPath = Left$(sPathIn, i)

End Function

Public Function getFileName(sFileIn As String) As String
' Essa função irá retornar apenas o nome do  arquivo de uma
' string que contenha o path e o nome do arquiva
Dim i As Integer

  For i = Len(sFileIn) To 1 Step -1
     If InStr("\", Mid$(sFileIn, i, 1)) Then Exit For
  Next

  getFileName = Left(Mid$(sFileIn, i + 1, Len(sFileIn) - i), Len(Mid$(sFileIn, i + 1, Len(sFileIn) - i)) - 4)

End Function

Public Function getFileExt(sFileIn As String) As String
' Essa função irá retornar apenas a extensão do  arquivo de uma
' string que contenha o path e o nome do arquiva
Dim i As Integer

  For i = Len(sFileIn) To 1 Step -1
     If InStr("\", Mid$(sFileIn, i, 1)) Then Exit For
  Next

  getFileExt = Right(Mid$(sFileIn, i + 1, Len(sFileIn) - i), 4)

End Function

Public Function getFileNameAndExt(sFileIn As String) As String
' Essa função irá retornar apenas o nome do  arquivo de uma
' string que contenha o path e o nome do arquiva
Dim i As Integer

    For i = Len(sFileIn) To 1 Step -1
       If InStr("\", Mid$(sFileIn, i, 1)) Then Exit For
    Next
    
    getFileNameAndExt = Mid$(sFileIn, i + 1, Len(sFileIn) - i)

End Function

Public Function pathDesktopAddress() As String
    pathDesktopAddress = CreateObject("WScript.Shell").SpecialFolders("Desktop") & "\"
End Function

Public Function pathWorkSheetAddress() As String
    pathWorkSheetAddress = ActiveWorkbook.Path & "\"
End Function

Public Function pathWorkbookFullName() As String
    pathWorkbookFullName = ActiveWorkbook.FullName
End Function



'Sub AddModuleToProject()
'Dim module As VBComponent
'Dim CreateModule
'
'Set module = ActiveWorkbook.VBProject.VBComponents.Add(vbext_ct_StdModule)
'module.Name = "MyModule"
'
'module.CodeModule.AddFromString "public sub test()" & vbNewLine & _
'                                "    'dosomething" & vbNewLine & _
'                                "end sub"
'Set CreateModule = module
'
'End Sub

