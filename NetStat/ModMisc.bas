Attribute VB_Name = "ModMisc"
Option Explicit
Public Declare Function RegisterWindowMessage Lib "user32" Alias "RegisterWindowMessageA" (ByVal lpString As String) As Long
Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (dest As Any, Src As Any, ByVal cb&)

Private Declare Function GetShortPathName Lib "kernel32" Alias "GetShortPathNameA" (ByVal lpszLongPath As String, ByVal lpszShortPath As String, ByVal cchBuffer As Long) As Long
Private Declare Function InternetGetConnectedState Lib "wininet.dll" (ByRef lpdwFlags As Long, ByVal dwReserved As Long) As Long
Private Declare Function IsTextUnicode Lib "advapi32" (lpBuffer As Any, ByVal cb As Long, lpi As Long) As Long
Private Declare Function LockWindowUpdate Lib "user32" (ByVal hwndLock As Long) As Long
Private Declare Function PathMatchSpecW Lib "shlwapi" (ByVal pszFileParam As Long, ByVal pszSpec As Long) As Long
Private Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hwndInsertAfter As Long, ByVal X As Long, Y, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Private Declare Function SHGetShortPathName Lib "shell32" Alias "#92" (ByVal szPath As String) As Long

'I shortened these from IS_TEXT_UNICODE_* to ITU_
Private Const ITU_REVERSE_STATISTICS As Long = &H20
Private Const ITU_ASCII16           As Long = &H1
Private Const ITU_REVERSE_ASCII16   As Long = &H10
Private Const ITU_STATISTICS        As Long = &H2
Private Const ITU_CONTROLS          As Long = &H4
Private Const ITU_REVERSE_CONTROLS  As Long = &H40
Private Const ITU_SIGNATURE         As Long = &H8
Private Const ITU_REVERSE_SIGNATURE As Long = &H80
Private Const ITU_ILLEGAL_CHARS     As Long = &H100
Private Const ITU_ODD_LENGTH        As Long = &H200
Private Const ITU_DBCS_LEADBYTE     As Long = &H400
Private Const ITU_NULL_BYTES        As Long = &H1000
Private Const ITU_UNICODE_MASK      As Long = &HF
Private Const ITU_REVERSE_MASK      As Long = &HF0
Private Const ITU_NOT_UNICODE_MASK  As Long = &HF00
Private Const ITU_NOT_ASCII_MASK    As Long = &HF000

Global g_TcpConnections()           As tConnectionType
Global g_UdpConnections()           As tConnectionType
Global g_bXPTable                   As Boolean
Global g_bIsWinNT                   As Boolean
Global g_sShell32Path               As String
Global g_aProgramDescriptions       As Dictionary
Global g_aPrograms()                As tProgram
Global g_bWorking                   As Boolean
Global g_bStopAll                   As Boolean
Private Const MAX_PATH              As Integer = 260
Private Type OSVERSIONINFO
    dwOSVersionInfoSize             As Long
    dwMajorVersion                  As Long
    dwMinorVersion                  As Long
    dwBuildNumber                   As Long
    dwPlatformId                    As Long
    szCSDVersion                    As String * 128
End Type
Type tChecking
    sName                           As String
    bBlocked                        As Boolean
End Type
Public Function GetShortPath(Path As String) As String
    Dim sRet                        As String * 260
    Dim lRet                        As Long
    lRet = GetShortPathName(Path, sRet, 260)
    If lRet = 0 Then
        GetShortPath = Path
    Else
        GetShortPath = Mid(sRet, 1, lRet)
    End If
End Function
Public Function SHGetShortPath(sPathIn As String) As String
    Dim sPathOut                    As String
    sPathOut = MakeMaxPath(sPathIn)
    sPathOut = CheckString(sPathOut)
    SHGetShortPathName sPathOut
    SHGetShortPath = GetStrFromBuffer(sPathOut)
End Function
Public Sub Main()
    g_bXPTable = CallApiByName("iphlpapi.dll", "AllocateAndGetTcpExTableFromStack") 'Check existance of this function.
        'The above dll ships with every version of windows since Windows 95.
        'BUT the above function is XP dependant to my knowledge.
        'That is why I have included two methods of performing the Tcp/Udp Tables
    g_bIsWinNT = IsWinNT
    FrmMain.Show
End Sub
Public Function FixPath(sPath As String, Optional AddSlash As Boolean = True) As String
    Select Case AddSlash
        Case Is = True
            FixPath = sPath & IIf(Right(sPath, 1) <> "\", "\", "") 'Add a \ if it doesn't exist.
        Case Is = False
            FixPath = IIf(Right(sPath, 1) = "\", Left$(sPath, Len(sPath) - 1), sPath) 'Add a \ if it doesn't exist.
    End Select
End Function
Public Function FindProgram(sFileName As String) As String
    If InStr(sFileName, ".") = 0 Then sFileName = sFileName & ".dll" 'Add .dll to the end to check existance.
    Dim sTmp                        As String
    Dim sArr()                      As String
    Dim X                           As Integer
    sTmp = FixPath(App.Path) & sFileName 'Add \ and FileName to app.path
    If Len(Dir(sTmp)) > 0 Then 'If it's contained in the App Directory, then that is the one windows used.
        FindProgram = sTmp
        Exit Function
    End If
    sArr = Split(Environ("PATH"), ";") 'Loops through all Environment Paths
    For X = 0 To UBound(sArr)
        sTmp = FixPath(sArr(X)) & sFileName
        If Len(Dir(sTmp)) > 0 Then 'The first one we find, is also the first one Windows finds *I think*
            FindProgram = sTmp
            Exit Function
        End If
    Next
End Function
Public Sub LockControl(Cntrl As Control, Optional UnLockIt As Boolean = False)
    On Error Resume Next
    'This stops a control from updating while you work with it.
    'Saves ALOT of time when adding lists etc.
    'When I say alot, I really mean alot lol.
    Select Case UnLockIt
        Case True
            Call LockWindowUpdate(0)
            Cntrl.Refresh
        Case False
            Call LockWindowUpdate(Cntrl.hwnd)
    End Select
End Sub
Public Function IsConnected() As Boolean
    IsConnected = InternetGetConnectedState(0&, 0&)
End Function
Public Function MatchSpec(ByVal sText As String, ByVal sSearch As String) As Boolean
    On Error Resume Next
    'I know, I know, this is for file searches, well, that 'was' it's purpose before I found it anyway =D
    'This little euphoric function basically does a Regular Expression match on the sText using sSearch as the expression.
    'Don't complain about the method if it gets the job done quickly and easy ;)
    MatchSpec = PathMatchSpecW(StrPtr(sText), StrPtr(sSearch))
End Function
Public Sub MakeOntop(hwnd As Long, Optional bNotTop As Boolean = False)
    Const HWND_TOPMOST              As Long = -1
    Const HWND_NOTOPMOST            As Long = -2
    Const SWP_NOMOVE                As Long = &H2
    Const SWP_NOSIZE                As Long = &H1
    Const lFlags                    As Long = SWP_NOMOVE Or SWP_NOSIZE
    If bNotTop = False Then
        Call SetWindowPos(hwnd, HWND_TOPMOST, 0&, 0&, 0&, 0&, lFlags) 'Set Top Most
    Else
        Call SetWindowPos(hwnd, HWND_NOTOPMOST, 0&, 0&, 0&, 0&, lFlags) 'Unset Top Most
    End If
End Sub
Public Function MakeByte(HiNibble As Byte, LoNibble As Byte) As Integer
    MakeByte = LoNibble
    MakeByte = MakeByte Or (HiNibble * 2)
End Function
Public Function LoNibble(bytValue As Byte) As Byte
    LoNibble = (bytValue And &HF)
End Function
Public Function HiNibble(bytValue As Byte) As Byte
    HiNibble = (bytValue And &HF0) \ 16
End Function
Private Function MakeMaxPath(ByVal sPath As String) As String
    'Terminates sPath w/ null chars making
    'the return string MAX_PATH chars long.
    MakeMaxPath = sPath & String$(MAX_PATH - Len(sPath), 0)
End Function
Private Function CheckString(Msg As String) As String
    If g_bIsWinNT Then
        CheckString = StrConv(Msg, vbUnicode)
    Else
        CheckString = Msg
    End If
End Function
Private Function GetStrFromBuffer(szStr As String) As String
    'Returns string before first null
    'char encountered (if any) from either an ANSII or
    'Unicode string buffer.
    If IsUnicodeStr(szStr) Then szStr = StrConv(szStr, vbFromUnicode)
    If InStr(szStr, vbNullChar) Then
        GetStrFromBuffer = Left$(szStr, InStr(szStr, vbNullChar) - 1)
    Else
        GetStrFromBuffer = szStr
    End If
End Function
Public Function IsUnicodeStr(sBuffer As String) As Boolean
    'Returns True if sBuffer evaluates to
    'a Unicode string
    Dim dwRtnFlags                  As Long
    dwRtnFlags = ITU_UNICODE_MASK
    IsUnicodeStr = IsTextUnicode(ByVal sBuffer, Len(sBuffer), dwRtnFlags)
End Function
Public Function InIDE() As Boolean
    On Error Resume Next
    Err.Clear
    Debug.Print 1 / 0
    If Err.Number > 0 Then
        InIDE = True
    Else
        InIDE = False
    End If
End Function
Public Function FindPrograms(ProgLocation As String) As Integer
    Dim X                           As Integer
    Dim sPath                       As String
    Dim bShort                      As Boolean
    Dim sShort                      As String
    sPath = LCase(GetShortPath(ProgLocation))
    For X = 1 To FrmMain.lstPrograms.ListItems.Count
        sShort = LCase(GetShortPath(FrmMain.lstPrograms.ListItems(X).key))
        If Len(sShort) > 0 Then
            If StrComp(sShort, ProgLocation, vbTextCompare) = 0 Then
                FindPrograms = X
                Exit Function
            End If
        End If
        If StrComp(FrmMain.lstPrograms.ListItems(X).key, ProgLocation, vbTextCompare) = 0 Then
            FindPrograms = X
            Exit Function
        End If
    Next
    FindPrograms = -1
End Function
