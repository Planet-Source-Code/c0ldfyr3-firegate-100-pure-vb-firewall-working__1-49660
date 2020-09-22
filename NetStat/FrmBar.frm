VERSION 5.00
Begin VB.Form FrmBar 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   510
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6750
   Icon            =   "FrmBar.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   510
   ScaleWidth      =   6750
   ShowInTaskbar   =   0   'False
   Begin VB.Timer tmrClose 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   4560
      Top             =   120
   End
   Begin VB.Timer tmrOpen 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   3960
      Top             =   120
   End
   Begin VB.TextBox txtAlert 
      Alignment       =   2  'Center
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   400
      Left            =   600
      MultiLine       =   -1  'True
      TabIndex        =   1
      Top             =   60
      Width           =   5895
   End
   Begin VB.Timer tmrMessage 
      Enabled         =   0   'False
      Interval        =   10000
      Left            =   4200
      Top             =   50
   End
   Begin VB.Timer TmrHide 
      Interval        =   4000
      Left            =   3720
      Top             =   50
   End
   Begin VB.Image ImgImage 
      Height          =   375
      Left            =   0
      MouseIcon       =   "FrmBar.frx":0CCA
      MousePointer    =   99  'Custom
      Top             =   60
      Width           =   315
   End
   Begin VB.Label lblAlert 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   480
      TabIndex        =   0
      Top             =   60
      Width           =   6135
   End
   Begin VB.Image ImgShape 
      Height          =   510
      Left            =   0
      Picture         =   "FrmBar.frx":19CC
      Top             =   0
      Width           =   6735
   End
End
Attribute VB_Name = "FrmBar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Private Type POINTAPI
    X                               As Long
    Y                               As Long
End Type
Dim Hidden                          As Boolean
Dim lThread                         As Long
Dim lOpen(2)                        As Integer
Dim lClose(2)                       As Integer
Private Sub Form_Load()
    Me.Left = Screen.Width - 187
    If Reshaping Then
        ImgShape = LoadPicture(Reshape_Map)
        SavePicture ImgShape.Picture, App.Path & "\Tmp.tmp"
        Face = CreateRegionFromFile(Me, ImgShape, App.Path & "\Tmp.tmp", RGB(0, 255, 0))
        SetWindowRgn Me.hwnd, Face, True
        HideMe
        Me.Visible = True
        Reshaping = False
    Else
        SavePicture ImgShape.Picture, App.Path & "\Tmp.tmp"
        Face = CreateRegionFromFile(Me, ImgShape, App.Path & "\Tmp.tmp", RGB(0, 255, 0))
        SetWindowRgn Me.hwnd, Face, True
        HideMe
        Me.Visible = True
    End If
End Sub
Private Sub HideMe()
    ImgImage.ToolTipText = "Double click to open"
    DoEvents
    DoClose
    Me.Left = Screen.Width - 187
    Hidden = True
End Sub
Public Sub ShowAlert(Optional Message As String)
    If Message <> "" Then txtAlert.Text = Message
    If Len(Message) > 45 Then
        txtAlert.Height = 400 / Screen.TwipsPerPixelX
        txtAlert.Top = 60 / Screen.TwipsPerPixelX
    Else
        txtAlert.Height = 200 / Screen.TwipsPerPixelX
        txtAlert.Top = (60 / Screen.TwipsPerPixelX) * 2
    End If
    If Hidden = False Then Exit Sub
    tmrMessage.Enabled = False
    TmrHide.Enabled = False
    ImgImage.ToolTipText = "Double click to close"
    Me.Left = Screen.Width - 187
    DoOpen
    Me.Left = Screen.Width - Me.Width
    Hidden = False
    TmrHide.Enabled = True
    tmrMessage.Enabled = True
End Sub
Public Sub DoOpen()
    lOpen(2) = -1
    lOpen(0) = Me.Left
    lOpen(1) = (Screen.Width - Me.Width)
    tmrOpen.Enabled = True
End Sub
Public Sub DoClose()
    lClose(2) = -1
    lClose(0) = Me.Left
    lClose(1) = (Screen.Width - 187)
    tmrClose.Enabled = True
End Sub
Private Sub ImgImage_DblClick()
    If Hidden Then ShowAlert
End Sub
Private Sub ImgImage_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = vbLeftButton Then
        ReleaseCapture
        SendMessage hwnd, WM_NCLBUTTONDOWN, HTCAPTION, ByVal 0&
    End If
    If Hidden Then
        ShowAlert
    End If
End Sub
Private Sub tmrClose_Timer()
    tmrOpen.Enabled = False
    If lClose(2) <= 0 Then lClose(2) = lClose(0)
    lClose(2) = lClose(2) + 100
    Me.Left = lClose(2)
    Me.Refresh
    If lClose(2) >= lClose(1) Then lClose(2) = -1: tmrClose.Enabled = False
End Sub
Private Sub TmrHide_Timer()
    If Not Hidden Then HideMe
End Sub
Private Sub tmrMessage_Timer()
    txtAlert.Text = "No Alert"
End Sub
Private Sub tmrOpen_Timer()
    Static X                        As Integer
    tmrClose.Enabled = False
    If lOpen(2) <= 0 Then lOpen(2) = lOpen(0)
    lOpen(2) = lOpen(2) - 100
    Me.Left = lOpen(2)
    Me.Refresh
    If lOpen(2) <= lOpen(1) Then lOpen(2) = -1: tmrOpen.Enabled = False
End Sub
Private Sub txtAlert_GotFocus()
    On Error Resume Next
    Me.SetFocus
End Sub
