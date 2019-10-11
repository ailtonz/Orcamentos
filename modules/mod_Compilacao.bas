Attribute VB_Name = "mod_Compilacao"
Option Explicit

'Declare API
#If VBA7 Then
    Declare PtrSafe Function GetKeyState Lib "USER32" (ByVal vKey As Long) As Integer
#Else
    Declare Function GetKeyState Lib "USER32" (ByVal vKey As Long) As Integer
#End If


#If VBA7 Then

    Declare PtrSafe Function OpenProcess Lib "kernel32" (ByVal dwDesiredAccess As LongPtr, ByVal bInheritHandle As LongPtr, ByVal dwProcessId As LongPtr) As LongPtr
    
    Declare PtrSafe Function GetTickCount64 Lib "kernel32" () As LongPtr

    Declare PtrSafe Function GetExitCodeProcess Lib "kernel32" (ByVal hProcess As LongPtr, lpExitCode As LongPtr) As LongPtr
    
    Declare PtrSafe Function InternetGetConnectedState Lib "wininet.dll" (ByRef dwflags As LongPtr, ByVal dwReserved As LongPtr) As LongPtr

#Else

    Declare Function OpenProcess Lib "kernel32" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwProcessId As Long) As Long

    Declare Function GetExitCodeProcess Lib "kernel32" (ByVal hProcess As Long, lpExitCode As Long) As Long
        
    Declare Function InternetGetConnectedState Lib "wininet.dll" (ByRef dwflags As Long, ByVal dwReserved As Long) As Long


#End If

Public Const PROCESS_QUERY_INFORMATION = &H400
Public Const STILL_ACTIVE = &H103


Public Sub ShellAndWait(ByVal PathName As String, Optional WindowState)

#If VBA7 Then
    Dim hProg As LongPtr
    Dim hProcess As LongPtr, ExitCode As LongPtr
#Else
    Dim hProg As Long
    Dim hProcess As Long, ExitCode As Long
#End If

    'fill in the missing parameter and execute the program
    If IsMissing(WindowState) Then WindowState = 1
        hProg = Shell(PathName, WindowState)
        'hProg is a "process ID under Win32. To get the process handle:
        hProcess = OpenProcess(PROCESS_QUERY_INFORMATION, False, hProg)
    Do
        'populate Exitcode variable
        GetExitCodeProcess hProcess, ExitCode
        DoEvents
    Loop While ExitCode = STILL_ACTIVE
    
End Sub


Function bIsBookOpen(ByRef szBookName As String) As Boolean
    On Error Resume Next
    bIsBookOpen = Not (Application.Workbooks(szBookName) Is Nothing)
End Function



