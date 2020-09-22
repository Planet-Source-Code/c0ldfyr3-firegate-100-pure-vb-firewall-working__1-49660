Attribute VB_Name = "ModProcs"
Option Explicit
Private Declare Function CloseHandle Lib "kernel32.dll" (ByVal Handle As Long) As Long
Private Declare Function CreateToolhelp32Snapshot Lib "kernel32" (ByVal dwFlags As Long, ByVal th32ProcessID As Long) As Long
Private Declare Function EnumProcesses Lib "psapi.dll" (ByRef lpidProcess As Long, ByVal cb As Long, ByRef cbNeeded As Long) As Long
Private Declare Function EnumProcessModules Lib "psapi.dll" (ByVal hProcess As Long, ByRef lphModule As Long, ByVal cb As Long, ByRef cbNeeded As Long) As Long
Private Declare Function GetCurrentProcessId Lib "kernel32" () As Long
Private Declare Function GetCurrentProcess Lib "kernel32" () As Long
Private Declare Function GetModuleFileNameExA Lib "psapi.dll" (ByVal hProcess As Long, ByVal hModule As Long, ByVal ModuleName As String, ByVal nSize As Long) As Long
Private Declare Function LookupAccountSid Lib "advapi32.dll" Alias "LookupAccountSidA" (ByVal lpSystemName As String, ByVal sID As Long, ByVal Name As String, cbName As Long, ByVal ReferencedDomainName As String, cbReferencedDomainName As Long, peUse As Integer) As Long
Private Declare Function OpenProcess Lib "kernel32.dll" (ByVal dwDesiredAccessas As Long, ByVal bInheritHandle As Long, ByVal dwProcId As Long) As Long
Private Declare Function Process32First Lib "kernel32" (ByVal hSnapShot As Long, lppe As PROCESSENTRY32) As Long
Private Declare Function Process32Next Lib "kernel32" (ByVal hSnapShot As Long, lppe As PROCESSENTRY32) As Long
Private Declare Function TerminateProcess Lib "kernel32" (ByVal ApphProcess As Long, ByVal uExitCode As Long) As Long
Private Declare Function WTSEnumerateProcesses Lib "wtsapi32.dll" Alias "WTSEnumerateProcessesA" (ByVal hServer As Long, ByVal Reserved As Long, ByVal Version As Long, ByRef ppProcessInfo As Long, ByRef pCount As Long) As Long
Private Declare Sub WTSFreeMemory Lib "wtsapi32.dll" (ByVal pMemory As Long)
Private Const PROCESS_QUERY_INFORMATION As Long = 1024
Private Const WTS_CURRENT_SERVER_HANDLE As Long = 0&
Private Const hNull                 As Long = 0
Private Const TH32CS_SNAPPROCESS    As Long = &H2&
Private Const PROCESS_VM_READ       As Long = 16
Private Const MAX_PATH              As Long = 260
Private Type PROCESSENTRY32
    dwSize                          As Long
    cntUsage                        As Long
    th32ProcessID                   As Long
    th32DefaultHeapID               As Long
    th32ModuleID                    As Long
    cntThreads                      As Long
    th32ParentProcessID             As Long
    pcPriClassBase                  As Long
    dwFlags                         As Long
    szExeFile                       As String * 260
End Type
Private Type WTS_PROCESS_INFO
    SessionID                       As Long
    ProcessID                       As Long
    pProcessName                    As Long
    pUserSid                        As Long
End Type
Private g_aProcessTable             As Dictionary
Private g_aProcessUIDs              As Dictionary
Public Sub Proc_Startup()
    Set g_aProcessTable = New Dictionary
    Set g_aProcessUIDs = New Dictionary
End Sub
Public Sub Proc_Cleanup()
    Set g_aProcessTable = Nothing
    Set g_aProcessUIDs = Nothing
End Sub
Public Function GetProcs(Optional kProc As Long = -1) As Boolean
    On Error Resume Next
    Dim lProcess                    As Long
    Dim lExitCode                   As Long
    Dim f                           As Long
    Dim sName                       As String
    Dim hSnap                       As Long
    Dim Proc                        As PROCESSENTRY32
    Dim cb                          As Long
    Dim cbNeeded                    As Long
    Dim NumElements                 As Long
    Dim cbNeeded2                   As Long
    Dim Modules(1 To 200)           As Long
    Dim lRet                        As Long
    Dim ModuleName                  As String
    Dim nSize                       As Long
    Dim hProcess                    As Long
    Dim i                           As Long
    Dim tmp                         As String
    Dim X                           As Long
    Dim sArr()                      As String
    Dim p_aProcessTable             As Dictionary
    Set p_aProcessTable = New Dictionary
    Select Case GetVersion()
        Case 1 'Windows 9x/Me
            hSnap = CreateToolhelp32Snapshot(TH32CS_SNAPPROCESS, 0) 'Get all process's
                'The above function works on all windows versions
                'But only returns the exe name on Windows NT.
                'On Win 9x/Me it returns full path.
            If hSnap = hNull Then
                GetProcs = False
                Exit Function
            End If
            Proc.dwSize = Len(Proc)
            f = Process32First(hSnap, Proc) 'Process First
            Do While f 'Loop through all process's
                If cGetInputState(QS_ALLEVENTS) <> 0 Then DoEvents
                sName = StrZToStr(Proc.szExeFile) 'Get Exe Path
                If kProc = Proc.th32ProcessID Then 'If we want to kill a process....
                    lProcess = OpenProcess(1, False, Proc.th32ProcessID) 'Open it for access
                    TerminateProcess lProcess, lExitCode 'Terminate
                    CloseHandle lProcess 'Close the open handle.
                Else
                    p_aProcessTable.Add Proc.th32ProcessID, sName 'Add the process id/exe name to our array.
                End If
                f = Process32Next(hSnap, Proc) 'Process Next
            Loop
        Case 2 'Windows NT/2K/XP
            cb = 8
            cbNeeded = 96
            Do While cb <= cbNeeded
                cb = cb * 2
                ReDim ProcessIDs(cb / 4) As Long
                lRet = EnumProcesses(ProcessIDs(1), cb, cbNeeded)  'Enumerate Process's
            Loop
            NumElements = cbNeeded / 4
            For i = 1 To NumElements 'Loop through all process's
                If cGetInputState(QS_ALLEVENTS) <> 0 Then DoEvents
                hProcess = OpenProcess(PROCESS_QUERY_INFORMATION Or PROCESS_VM_READ, 0, ProcessIDs(i)) 'Open process for access.
                If hProcess <> 0 Then
                    lRet = EnumProcessModules(hProcess, Modules(1), 200, cbNeeded2) 'Enumerate it's loaded modules.
                    If lRet <> 0 Then
                        ModuleName = Space(MAX_PATH)
                        nSize = 500
                        lRet = GetModuleFileNameExA(hProcess, Modules(1), ModuleName, nSize) 'Get the EXE Path
                        tmp = Left(ModuleName, lRet)
                        If kProc = ProcessIDs(i) Then 'If we want to kill a process....
                            lProcess = OpenProcess(1, False, ProcessIDs(i)) 'Open it
                            TerminateProcess lProcess, lExitCode 'Kill it
                            CloseHandle lProcess 'Close it
                        Else
                            p_aProcessTable.Add ProcessIDs(i), tmp 'Add the process
                            tmp = ""
                        End If
                    End If
                End If
                lRet = CloseHandle(hProcess) 'Close the open process handle.
            Next
        Case Else
            GetProcs = False
            Exit Function
    End Select
    Set g_aProcessTable = p_aProcessTable
    Set p_aProcessTable = Nothing
    GetProcs = True
End Function
Private Function StrZToStr(sText As String) As String
    StrZToStr = Left$(sText, Len(sText) - 1) 'Remove last character
End Function
Public Sub LoadProcessUserIDs()
    Dim udtProcessInfo              As WTS_PROCESS_INFO
    Dim lRet                        As Long
    Dim lCount                      As Long
    Dim X                           As Integer
    Dim lP                          As Long
    Dim lBuffer                     As Long
    Dim lPlace                      As Long
    Dim p_aProcessUIDs              As Dictionary
    Set p_aProcessUIDs = New Dictionary
    lRet = WTSEnumerateProcesses(WTS_CURRENT_SERVER_HANDLE, 0&, 1, lBuffer, lCount) 'Get the pointer to the Enumeration.
    If lRet Then
        lPlace = lBuffer
        For X = 1 To lCount
            If cGetInputState(QS_ALLEVENTS) <> 0 Then DoEvents
            CopyMemory udtProcessInfo, ByVal lBuffer, LenB(udtProcessInfo) 'Get each type struct.
            p_aProcessUIDs.Add udtProcessInfo.ProcessID, GetUserName(udtProcessInfo.pUserSid)  'Add the user name for each process.
            lBuffer = lBuffer + LenB(udtProcessInfo) 'Add the place holder.
        Next
        WTSFreeMemory lPlace 'Free the pointer from memory.
    End If
    Set g_aProcessUIDs = p_aProcessUIDs
    Set p_aProcessUIDs = Nothing
End Sub
Private Function GetUserName(sID As Long) As String
    On Error Resume Next
    Dim sName                       As String
    Dim sDomain                     As String
    Dim iPos                        As Integer
    sName = String(255, 0)
    sDomain = String(255, 0)
    LookupAccountSid vbNullString, sID, sName, 255, sDomain, 255, 0
    sName = Left$(sDomain, InStr(sDomain, vbNullChar) - 1) & "\" & Left$(sName, InStr(sName, vbNullChar) - 1)
    iPos = InStr(1, sName, "\")
    If iPos > 0 Then sName = Mid(sName, iPos + 1)
    GetUserName = sName
End Function
Public Function Proc_UserName(lProcessID As Long) As String
    On Error GoTo ErrClear
    Proc_UserName = g_aProcessUIDs(lProcessID)
    Exit Function
ErrClear:
    Proc_UserName = ""
    Err.Clear
End Function
Public Function Proc_Path(lProcessID As Long) As String
    On Error GoTo ErrClear
    Proc_Path = g_aProcessTable(lProcessID)
    Exit Function
ErrClear:
    Proc_Path = ""
    Err.Clear
End Function
Public Function MyProcess() As Long
    MyProcess = GetCurrentProcessId
End Function
