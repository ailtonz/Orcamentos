Attribute VB_Name = "basInternet"
'Private Declare Function InternetGetConnectedState _
'   Lib "wininet.dll" (ByRef dwflags As Long, _
'   ByVal dwReserved As Long) As Long


#If VBA7 Then
    Private Const INTERNET_CONNECTION_MODEM As LongPtr = &H1
    Private Const INTERNET_CONNECTION_LAN As LongPtr = &H2
    Private Const INTERNET_CONNECTION_PROXY As LongPtr = &H4
    Private Const INTERNET_CONNECTION_OFFLINE As LongPtr = &H20
#Else
    Private Const INTERNET_CONNECTION_MODEM As Long = &H1
    Private Const INTERNET_CONNECTION_LAN As Long = &H2
    Private Const INTERNET_CONNECTION_PROXY As Long = &H4
    Private Const INTERNET_CONNECTION_OFFLINE As Long = &H20
#End If

Function IsInternetConnected() As Boolean

#If VBA7 Then
    Dim L As LongPtr
    Dim R As LongPtr
#Else
    Dim L As Long
    Dim R As Long
#End If

    R = InternetGetConnectedState(L, 0&)
    If R = 0 Then
        IsInternetConnected = False
    Else
        If R <= 4 Then
            IsInternetConnected = True
        Else
            IsInternetConnected = False
        End If
    End If
End Function
'You would call this in your code with something like
 
Sub AAA()
    If IsInternetConnected() = True Then
        ' connected
        MsgBox "OK"
    Else
        ' no connected
        MsgBox "Ñ.OK"
    End If
End Sub

