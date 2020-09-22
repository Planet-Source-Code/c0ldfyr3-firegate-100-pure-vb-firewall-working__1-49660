Attribute VB_Name = "ModTextBox"
Option Explicit
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Public Enum enTextBox
    ES_LEFT = &H0&
    ES_CENTER = &H1&
    ES_RIGHT = &H2&
    ES_MULTILINE = &H4&
    ES_UPPERCASE = &H8&
    ES_LOWERCASE = &H10&
    ES_PASSWORD = &H20&
    ES_AUTOVSCROLL = &H40&
    ES_AUTOHSCROLL = &H80&
    ES_NOHIDESEL = &H100&
    ES_OEMCONVERT = &H400&
    ES_READONLY = &H800&
    ES_WANTRETURN = &H1000&
    ES_NUMBER = &H2000&
End Enum
Private Const GWL_STYLE             As Long = (-16)
Public Sub MakeNumberOnly(TxtBox As TextBox)
    Static lStyle                   As Long
    lStyle = GetWindowLong(TxtBox.hwnd, GWL_STYLE)
    Call SetWindowLong(TxtBox.hwnd, GWL_STYLE, lStyle Or ES_NUMBER)
    TxtBox.Refresh
End Sub

