VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsWaitableTimer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private Declare Function CreateWaitableTimer Lib "kernel32" Alias "CreateWaitableTimerA" (ByVal lpSemaphoreAttributes As Long, ByVal bManualReset As Long, ByVal lpName As String) As Long
Private Declare Function OpenWaitableTimer Lib "kernel32" Alias "OpenWaitableTimerA" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal lpName As String) As Long
Private Declare Function SetWaitableTimer Lib "kernel32" (ByVal hTimer As Long, lpDueTime As FILETIME, ByVal lPeriod As Long, ByVal pfnCompletionRoutine As Long, ByVal lpArgToCompletionRoutine As Long, ByVal fResume As Long) As Long
Private Declare Function CancelWaitableTimer Lib "kernel32" (ByVal hTimer As Long)
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Private Declare Function WaitForSingleObject Lib "kernel32" (ByVal hHandle As Long, ByVal dwMilliseconds As Long) As Long
Private Declare Function MsgWaitForMultipleObjects Lib "user32" (ByVal nCount As Long, pHandles As Long, ByVal fWaitAll As Long, ByVal dwMilliseconds As Long, ByVal dwWakeMask As Long) As Long
Private Const WAIT_ABANDONED        As Long = &H80&
Private Const WAIT_ABANDONED_0      As Long = &H80&
Private Const WAIT_FAILED           As Long = -1&
Private Const WAIT_IO_COMPLETION    As Long = &HC0&
Private Const WAIT_OBJECT_0         As Long = 0
Private Const WAIT_OBJECT_1         As Long = 1
Private Const WAIT_TIMEOUT          As Long = &H102&
Private Const INFINITE              As Long = &HFFFF
Private Const ERROR_ALREADY_EXISTS  As Long = 183&
Private Const QS_HOTKEY             As Long = &H80
Private Const QS_KEY                As Long = &H1
Private Const QS_MOUSEBUTTON        As Long = &H4
Private Const QS_MOUSEMOVE          As Long = &H2
Private Const QS_PAINT              As Long = &H20
Private Const QS_POSTMESSAGE        As Long = &H8
Private Const QS_SENDMESSAGE        As Long = &H40
Private Const QS_TIMER              As Long = &H10
Private Const QS_MOUSE              As Long = (QS_MOUSEMOVE Or QS_MOUSEBUTTON)
Private Const QS_INPUT              As Long = (QS_MOUSE Or QS_KEY)
Private Const QS_ALLEVENTS          As Long = (QS_INPUT Or QS_POSTMESSAGE Or QS_TIMER Or QS_PAINT Or QS_HOTKEY)
Private Const QS_ALLINPUT           As Long = (QS_SENDMESSAGE Or QS_PAINT Or QS_TIMER Or QS_POSTMESSAGE Or QS_MOUSEBUTTON Or QS_MOUSEMOVE Or QS_HOTKEY Or QS_KEY)
Private Const Units                 As Double = 4294967296#
Private Const MAX_LONG              As Double = -2147483648#
Private mlTimer                     As Long
Private Type FILETIME
    dwLowDateTime                   As Long
    dwHighDateTime                  As Long
End Type
Private Sub Class_Terminate()
    If mlTimer <> 0 Then CloseHandle mlTimer
End Sub
Public Sub Wait(MilliSeconds As Long)
    Dim ft                          As FILETIME
    Dim lBusy                       As Long
    Dim lRet                        As Long
    Dim dblDelay                    As Double
    Dim dblDelayLow                 As Double
    mlTimer = CreateWaitableTimer(0, True, App.EXEName & "Timer" & Format$(Now(), "NNSS"))
    If Err.LastDllError <> ERROR_ALREADY_EXISTS Then
        ft.dwLowDateTime = -1
        ft.dwHighDateTime = -1
        lRet = SetWaitableTimer(mlTimer, ft, 0, 0, 0, 0)
    End If
    dblDelay = CDbl(MilliSeconds) * 10000#
    ft.dwHighDateTime = -CLng(dblDelay / Units) - 1
    dblDelayLow = -Units * (dblDelay / Units - Fix(CStr(dblDelay / Units)))
    If dblDelayLow < MAX_LONG Then dblDelayLow = Units + dblDelayLow
    ft.dwLowDateTime = CLng(dblDelayLow)
    lRet = SetWaitableTimer(mlTimer, ft, 0, 0, 0, False)
    Do
        lBusy = MsgWaitForMultipleObjects(1, mlTimer, False, INFINITE, QS_ALLINPUT&)
        DoEvents
    Loop Until lBusy = WAIT_OBJECT_0
    CloseHandle mlTimer
    mlTimer = 0
End Sub
