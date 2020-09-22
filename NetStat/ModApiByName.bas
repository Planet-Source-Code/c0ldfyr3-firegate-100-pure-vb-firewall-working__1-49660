Attribute VB_Name = "ModApiByName"
Option Explicit
'This Module courtesy of M.A. Munim, http://www.munim.tk/
'He had posted this on a message board at http://experts-exchange.com
'Comments by me ;)
'Thanks !
Private Declare Function LoadLibrary Lib "kernel32" Alias "LoadLibraryA" (ByVal lpLibFileName As String) As Long
Private Declare Function GetProcAddress Lib "kernel32" (ByVal hModule As Long, ByVal lpProcName As String) As Long
Private Declare Function FreeLibrary Lib "kernel32" (ByVal hLibModule As Long) As Long
Public Function CallApiByName(libName As String, funcName As String) As Boolean
    On Error GoTo ErrClear 'Just incase an error occurs that I haven't thought about.
    Dim lLib                        As Long
    Dim i                           As Integer
    Dim sVer                        As String
    Dim mlngAddress                 As Long
    lLib = LoadLibrary(ByVal libName) 'Load the library.
    If lLib = 0 Then
        'If it doesn't exist.
        MsgBox libName & " not found." & vbCrLf & "Try finding it on google.ie, search for ""Download " & libName & """ and placing it in your " & SpecialFolder(WinSystem) & " folder.", vbCritical, App.Title
        CallApiByName = False
        Exit Function
    End If
    mlngAddress = GetProcAddress(lLib, ByVal funcName) 'Get function handle.
    If mlngAddress = 0 Then
        sVer = FindProgram(libName) 'Find the dll location.
        sVer = FileInfo(sVer, [File Version]) 'Get the version.
        MsgBox "Function entry point not found." & vbCrLf & "If you think you're version of windows should be supported by this application and you are getting this error, please e-mail c0ldfyr3@eliteprodigy.com with the following information." & vbCrLf & vbCrLf & "Library: " & libName & vbCrLf & "Function: " & funcName & vbCrLf & "Version: " & sVer, vbCritical
        CallApiByName = False
        GoTo ExitFunc
    End If
    CallApiByName = True
ExitFunc:
    If CallApiByName = False Then Call MsgBox(App.Title & " will now use toned down functions to perform the actions it requires.", vbInformation, App.Title)
    FreeLibrary lLib
    Exit Function
ErrClear:
    Call MsgBox("Error: #" & Err.Number & vbCrLf & Err.Description & "CallAPIByName()", vbExclamation, App.Title)
    Err.Clear
End Function
