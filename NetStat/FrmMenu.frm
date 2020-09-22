VERSION 5.00
Begin VB.Form FrmMenu 
   Caption         =   "Form1"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
End
Attribute VB_Name = "FrmMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim lMsg                        As Long
    If FrmMain.ScaleMode = vbPixels Then
        lMsg = X
    Else
        lMsg = X / Screen.TwipsPerPixelX
    End If
    Select Case lMsg
        Case WM_LBUTTONUP
            If FrmMain.Visible = False Then
                FrmMain.mnPopShow.Caption = "|Hide The Main Window|&Hide"
                FrmMain.WindowState = vbNormal
                FrmMain.Visible = True
                SetForegroundWindow FrmMain.hwnd
            End If
        Case WM_LBUTTONDBLCLK
        Case WM_RBUTTONUP
            SetForegroundWindow FrmMain.hwnd
            PopupMenu FrmMain.mnPopup
    End Select
    Exit Sub
Form_MouseMove_err:
End Sub
Private Sub Form_Unload(Cancel As Integer)
    Shell_NotifyIcon NIM_DELETE, nidProgramData
End Sub
