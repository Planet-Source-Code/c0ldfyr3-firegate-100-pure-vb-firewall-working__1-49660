Attribute VB_Name = "ModSysTray"
Option Explicit
Public Declare Function Shell_NotifyIcon Lib "shell32" Alias "Shell_NotifyIconA" (ByVal dwMessage As enm_NIM_Shell, pnid As NOTIFYICONDATA) As Boolean
Public Declare Function SetForegroundWindow Lib "user32" (ByVal hwnd As Long) As Long
Public nidProgramData           As NOTIFYICONDATA
Public Const WM_MOUSEISMOVING   As Long = &H200  'Mouse is moving
Public Const WM_LBUTTONDOWN     As Long = &H201  'Button down
Public Const WM_LBUTTONUP       As Long = &H202  'Button up
Public Const WM_LBUTTONDBLCLK   As Long = &H203  'Double-click
Public Const WM_RBUTTONDOWN     As Long = &H204  'Button down
Public Const WM_RBUTTONUP       As Long = &H205  'Button up
Public Const WM_RBUTTONDBLCLK   As Long = &H206  'Double-click
Public Const WM_SETHOTKEY       As Long = &H32
Public Const WM_MOUSEMOVE       As Long = &H200
Public Type NOTIFYICONDATA
    cbSize                      As Long
    hwnd                        As Long
    uId                         As Long
    uFlags                      As Long
    uCallbackMessage            As Long
    hIcon                       As Long
    szTip                       As String * 64
End Type
Public Enum enm_NIM_Shell
    NIM_ADD = &H0
    NIM_MODIFY = &H1
    NIM_DELETE = &H2
    NIF_MESSAGE = &H1
    NIF_ICON = &H2
    NIF_TIP = &H4
End Enum
'~~~~~~~Load Up~~~~~~~
'With nidProgramData
'             .cbSize = Len(nidProgramData)
'             .hWnd = Me.hWnd
'             .uId = vbNull
'             .uFlags = NIF_ICON Or NIF_TIP Or NIF_MESSAGE
'             .uCallbackMessage = WM_MOUSEMOVE
'             .hIcon = Me.Icon
'             .szTip = "Yahoo Auto Responce Bot" & vbNullChar
'End With
'Shell_NotifyIcon NIM_ADD, nidProgramData

'~~~~~~~Close It~~~~~~~
'Shell_NotifyIcon NIM_DELETE, nidProgramData

'~~~~~~Mouse Move~~~~~
'    On Error GoTo Form_MouseMove_err:
'    Dim Result, MSG As Long, I As Integer
'    If Me.ScaleMode = vbPixels Then
'        MSG = X
'    Else
'        MSG = X / Screen.TwipsPerPixelX
'    End If
'    Select Case MSG
'        Case WM_LBUTTONUP
'           SetForegroundWindow me.hwnd
'           Your Code Here
'        Case WM_LBUTTONDBLCLK
'           SetForegroundWindow me.hwnd
'           Your Code Here
'        Case WM_RBUTTONUP
'           SetForegroundWindow me.hwnd
'           Your Code Here
'    End Select
'    Exit Sub
'Form_MouseMove_err:



