Attribute VB_Name = "basConnection"
Sub testev2()
Dim cnn As ADODB.connection: Set cnn = New ADODB.connection
Dim strCnn As String: strCnn = "Provider=MSDASQL;Driver={MySQL ODBC 5.3 Ansi Driver};Server=186.202.152.40;Database=Ailton_springer;PORT=3306;UID=Ailto_springer;PWD=41L70N@@;Option=3;"

Call cnn.Open(strCnn)

If cnn.State = 1 Then
    MsgBox "ok"
Else
    MsgBox "n.ok"
End If



End Sub


Sub teste()

Dim connection As New ADODB.connection


loadBancos

    Set connection = OpenConnection(Banco(0))
    '' Is Connected
    If connection.State = 1 Then
        MsgBox "ok"
    Else
        MsgBox "n.ok"
    End If


End Sub


Public Function OpenConnection(strBanco As infBanco) As ADODB.connection
'' Build the connection string depending on the source
Dim connectionString As String
    
Select Case strBanco.strSource
    Case "Access"
        connectionString = "Provider=" & strBanco.strDriver & ";Data Source=" & strBanco.strDatabase
    Case "Access2003"
        connectionString = "Driver={" & strBanco.strDriver & "};Dbq=" & strBanco.strLocation & strBanco.strDatabase & ";Uid=" & strBanco.strUser & ";PWD=" & strBanco.strPassword & ""
    Case "SQLite"
        connectionString = "Driver={" & strBanco.strDriver & "};Database=" & strBanco.strDatabase
    Case "MySQL"
        connectionString = "Provider=MSDASQL;Driver={" & strBanco.strDriver & "};Server=" & strBanco.strLocation & ";Database=" & strBanco.strDatabase & ";PORT=" & strBanco.strPort & ";UID=" & strBanco.strUser & ";PWD=" & strBanco.strPassword & ";Option=3;"
    Case "PostgreSQL"
        connectionString = "Driver={" & strBanco.strDriver & "};Server=" & strBanco.strLocation & ";Database=" & strBanco.strDatabase & ";UID=" & strBanco.strUser & ";PWD=" & strBanco.strPassword
End Select

'' Create and open a new connection to the selected source
Set OpenConnection = New ADODB.connection
Call OpenConnection.Open(connectionString)
   
End Function
