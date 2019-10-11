VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmAnexosArquivos 
   Caption         =   "Anexos"
   ClientHeight    =   6300
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   10950
   OleObjectBlob   =   "frmAnexosArquivos.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmAnexosArquivos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub UserForm_Initialize()

Dim strBanco As String: strBanco = Range(BancoLocal)
Dim strControle As String: strControle = ActiveSheet.Name
Dim strVendedor As String: strVendedor = Range(GerenteDeContas).value

Dim strSQL_ANEXOS As String
strSQL_ANEXOS = "Select * from qryOrcamentosArquivosAnexos where Vendedor = '" & strVendedor & "' AND Controle = '" & strControle & "'"

ListBoxCarregar strBanco, Me, Me.lstAnexos.Name, "OBS_01", strSQL_ANEXOS
    
End Sub

Private Sub cmdNovo_Click()

Dim strBanco As String: strBanco = Range(BancoLocal)
Dim strControle As String: strControle = ActiveSheet.Name
Dim strVendedor As String: strVendedor = Range(GerenteDeContas).value
    
admArquivosAnexosCarregar strBanco, strControle, strVendedor

Dim strSQL_ANEXOS As String
strSQL_ANEXOS = "Select * from qryOrcamentosArquivosAnexos where Vendedor = '" & strVendedor & "' AND Controle = '" & strControle & "'"
ListBoxCarregar strBanco, Me, Me.lstAnexos.Name, "OBS_01", strSQL_ANEXOS
    
    
End Sub

Private Sub cmdExcluir_Click()

Dim strBanco As String: strBanco = Range(BancoLocal)
Dim strControle As String: strControle = ActiveSheet.Name
Dim strVendedor As String: strVendedor = Range(GerenteDeContas).value

Dim strSQL_ANEXOS As String

admOrcamentoExcluirAnexoArquivo strBanco, strVendedor, strControle, Me.lstAnexos.value
strSQL_ANEXOS = "Select * from qryOrcamentosArquivosAnexos where Vendedor = '" & strVendedor & "' AND Controle = '" & strControle & "'"
ListBoxCarregar strBanco, Me, Me.lstAnexos.Name, "OBS_01", strSQL_ANEXOS

End Sub

Private Sub lstAnexos_DblClick(ByVal Cancel As MSForms.ReturnBoolean)

If Me.lstAnexos.value <> "" Then
       
    'Set Dimension
    Dim objFSO
    Set objFSO = CreateObject("Scripting.FileSystemObject")
    
    If objFSO.FileExists(Me.lstAnexos.value) Then
        
        Dim obj
        Set obj = CreateObject("WScript.Shell")
        obj.Run Chr(34) & Me.lstAnexos.value & Chr(34)
    
    Else
        MsgBox "ATENÇÃO: Arquivo inexistente !", vbOKOnly + vbInformation, "Arquivo inexistente"
    End If
    
End If
        

End Sub

Private Sub lstAnexos_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)

Dim strBanco As String: strBanco = Range(BancoLocal)
Dim strControle As String: strControle = ActiveSheet.Name
Dim strVendedor As String: strVendedor = Range(GerenteDeContas).value

Dim strSQL_ANEXOS As String

Select Case KeyCode

    Case vbKeyDelete
    
        admOrcamentoExcluirAnexoArquivo strBanco, strVendedor, strControle, Me.lstAnexos.value
        strSQL_ANEXOS = "Select * from qryOrcamentosArquivosAnexos where Vendedor = '" & strVendedor & "' AND Controle = '" & strControle & "'"
        ListBoxCarregar strBanco, Me, Me.lstAnexos.Name, "OBS_01", strSQL_ANEXOS
        
    Case vbKeyReturn
    
        Dim obj
        Set obj = CreateObject("WScript.Shell")
        obj.Run Chr(34) & Me.lstAnexos.value & Chr(34)
        
End Select

End Sub





