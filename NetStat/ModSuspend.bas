Attribute VB_Name = "ModSuspend"
Option Explicit
'Thanks to Ananda Raja for this one, http://www.pscode.com/vb/scripts/ShowCode.asp?txtCodeId=34560&lngWId=1
Private Declare Function CloseHandle Lib "kernel32.dll" (ByVal Handle As Long) As Long
Private Declare Function CreateToolhelp32Snapshot Lib "kernel32.dll" (ByVal dwFlags As Long, ByVal th32ProcessID As Long) As Long
Private Declare Function GetWindowThreadProcessId Lib "user32" (ByVal hwnd As Long, lpdwProcessId As Long) As Long
Private Declare Function OpenProcess Lib "kernel32.dll" (ByVal dwDesiredAccessas As Long, ByVal bInheritHandle As Long, ByVal dwProcId As Long) As Long
Private Declare Function OpenThread Lib "kernel32" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Boolean, ByVal dwThreadID As Long) As Long
Private Declare Function ResumeThread Lib "kernel32" (ByVal hThread As Long) As Long
Private Declare Function SuspendThread Lib "kernel32" (ByVal hThread As Long) As Long
Private Declare Function TerminateProcess Lib "kernel32" (ByVal ApphProcess As Long, ByVal uExitCode As Long) As Long
Private Declare Function Thread32First Lib "kernel32.dll" (ByVal hSnapShot As Long, ByRef lpte As THREADENTRY32) As Boolean
Private Declare Function Thread32Next Lib "kernel32.dll" (ByVal hSnapShot As Long, ByRef lpte As THREADENTRY32) As Boolean
Private Type THREADENTRY32
    dwSize                          As Long
    cntUsage                        As Long
    th32ThreadID                    As Long
    th32OwnerProcessID              As Long
    tpBasePri                       As Long
    tpDeltaPri                      As Long
    dwFlags                         As Long
End Type
Private Const THREAD_DIRECT_IMPERSONATION As Long = &H200
Private Const THREAD_QUERY_INFORMATION As Long = &H40
Private Const STANDARD_RIGHTS_REQUIRED As Long = &HF0000
Private Const THREAD_SET_THREAD_TOKEN As Long = &H80
Private Const THREAD_SET_INFORMATION As Long = &H20
Private Const SYNCHRONIZE            As Long = &H100000
Private Const THREAD_TERMINATE      As Long = &H1
Private Const THREAD_SUSPEND_RESUME As Long = &H2
Private Const THREAD_GET_CONTEXT    As Long = &H8
Private Const THREAD_SET_CONTEXT    As Long = &H10
Private Const THREAD_IMPERSONATE    As Long = &H100
Private Const THREAD_ALL_ACCESS     As Long = (STANDARD_RIGHTS_REQUIRED Or SYNCHRONIZE Or &H3FF)
Private Const TH32CS_SNAPHEAPLIST   As Long = &H1
Private Const TH32CS_SNAPPROCESS    As Long = &H2
Private Const TH32CS_SNAPTHREAD     As Long = &H4
Private Const TH32CS_SNAPMODULE     As Long = &H8
Private Const TH32CS_SNAPALL        As Long = (TH32CS_SNAPHEAPLIST Or TH32CS_SNAPPROCESS Or TH32CS_SNAPTHREAD Or TH32CS_SNAPMODULE)
Private Const TH32CS_INHERIT        As Long = &H80000000
Private bSuspended                  As Boolean
Private lPauseProc                  As Long
Private aThreads                    As Dictionary
Public Function EnumThreads(ByVal lProcessID As Long) As Long
    'On Error GoTo VB_Error
    Dim THREADENTRY32               As THREADENTRY32
    Dim hSnapShot                   As Long
    Dim lThread                     As Long
    hSnapShot = CreateToolhelp32Snapshot(TH32CS_SNAPTHREAD, lProcessID)
    THREADENTRY32.dwSize = Len(THREADENTRY32)
    If Thread32First(hSnapShot, THREADENTRY32) = False Then
        EnumThreads = -1
        Exit Function
    Else
        If THREADENTRY32.th32OwnerProcessID = lProcessID Then AddThread THREADENTRY32.th32ThreadID
    End If
    Do
        If cGetInputState(QS_ALLEVENTS) <> 0 Then DoEvents
        If Thread32Next(hSnapShot, THREADENTRY32) = False Then
            Exit Do
        Else
            If THREADENTRY32.th32OwnerProcessID = lProcessID Then
                lThread = lThread + 1
                AddThread THREADENTRY32.th32ThreadID
            End If
        End If
    Loop
    Call CloseHandle(hSnapShot)
    EnumThreads = lThread
Exit Function
VB_Error:
Resume Next
End Function
Public Function KillProcess(ProcessID As Long) As Boolean
    Dim lProcess                    As Long
    Dim lExitCode                   As Long
    If ProcessID = MyProcess Then Exit Function
    lProcess = OpenProcess(1, False, ProcessID) 'Open it for access
    TerminateProcess lProcess, lExitCode 'Terminate
    CloseHandle lProcess 'Close the open handle.
    Set aThreads = Nothing
End Function
Public Function PauseProcess(ProcessID As Long) As Boolean
    Dim X                           As Long
    Dim hThread                     As Long
    If ProcessID = MyProcess Then Exit Function
    Set aThreads = New Dictionary 'A sweet little associative collection.
    X = EnumThreads(ProcessID)
    If X = -1 Then
        PauseProcess = False 'If no threads, then how can we suspend them O.o ?
        Exit Function
    End If
    For X = 0 To X - 1
        hThread = OpenThread(THREAD_ALL_ACCESS, False, aThreads.Item(aThreads.Keys(X)))  'Open the thread for access.
        Call SuspendThread(hThread)
    Next
    bSuspended = True
    PauseProcess = True
End Function
Public Function UnPauseProcess(ProcessID As Long) As Boolean
    Dim X                           As Long
    Dim hThread                     As Long
    If ProcessID = MyProcess Then Exit Function
    If bSuspended = False Then UnPauseProcess = False: Exit Function
    X = EnumThreads(ProcessID)
    For X = 0 To X - 1 'Loop through
        If cGetInputState(QS_ALLEVENTS) <> 0 Then DoEvents
        hThread = OpenThread(THREAD_SUSPEND_RESUME, False, aThreads.Item(aThreads.Keys(X))) 'Open thread for access.
        Call ResumeThread(hThread)
    Next
    Set aThreads = Nothing 'Destroy the Associative Array
    UnPauseProcess = True
End Function
Private Function AddThread(ThreadID As Long)
    On Error GoTo ErrClear 'Easiest fastest way I could think of adding threads without duplicates.
    aThreads.Add ThreadID, ThreadID
ErrClear:
    Err.Clear
End Function
