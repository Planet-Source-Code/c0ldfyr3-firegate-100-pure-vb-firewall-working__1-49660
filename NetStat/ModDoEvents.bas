Attribute VB_Name = "ModDoEvents"
Option Explicit
Public Enum eInputStates
    QS_HOTKEY = &H80
    QS_KEY = &H1
    QS_MOUSEBUTTON = &H4
    QS_MOUSEMOVE = &H2
    QS_PAINT = &H20
    QS_POSTMESSAGE = &H8
    QS_SENDMESSAGE = &H40
    QS_TIMER = &H10
    QS_ALLINPUT = (QS_SENDMESSAGE Or QS_PAINT Or QS_TIMER Or QS_POSTMESSAGE Or QS_MOUSEBUTTON Or QS_MOUSEMOVE Or QS_HOTKEY Or QS_KEY)
    QS_MOUSE = (QS_MOUSEMOVE Or QS_MOUSEBUTTON)
    QS_INPUT = (QS_MOUSE Or QS_KEY)
    QS_ALLEVENTS = (QS_INPUT Or QS_POSTMESSAGE Or QS_TIMER Or QS_PAINT Or QS_HOTKEY)
End Enum
Public Declare Function GetQueueStatus Lib "user32" (ByVal qsFlags As Long) As Long
Public Function cGetInputState(State As eInputStates)
    Dim qsRet                   As Long
    qsRet = GetQueueStatus(State)
    cGetInputState = qsRet
End Function
