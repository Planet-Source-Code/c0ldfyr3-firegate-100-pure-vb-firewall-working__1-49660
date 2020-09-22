VERSION 5.00
Begin VB.Form frmAlert 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Fire Gate"
   ClientHeight    =   3255
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7545
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmAlert.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3255
   ScaleWidth      =   7545
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Pic32 
      BorderStyle     =   0  'None
      Height          =   480
      Left            =   120
      ScaleHeight     =   32
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   32
      TabIndex        =   4
      Top             =   600
      Visible         =   0   'False
      Width           =   480
   End
   Begin prjNetStat.lvButtons_H cmdYes 
      Height          =   375
      Left            =   5040
      TabIndex        =   2
      Top             =   2760
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   661
      Caption         =   "Yes"
      CapAlign        =   2
      BackStyle       =   3
      Shape           =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      cGradient       =   0
      Mode            =   0
      Value           =   0   'False
      cBack           =   -2147483633
   End
   Begin VB.CheckBox chkRemember 
      Caption         =   "Remeber my answer, and do not ask me again for this application."
      Height          =   375
      Left            =   240
      TabIndex        =   1
      Top             =   2760
      Width           =   2775
   End
   Begin prjNetStat.lvButtons_H cmdNo 
      Height          =   375
      Left            =   6000
      TabIndex        =   3
      Top             =   2760
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   661
      Caption         =   "No"
      CapAlign        =   2
      BackStyle       =   3
      Shape           =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      cGradient       =   0
      Mode            =   0
      Value           =   0   'False
      cBack           =   -2147483633
   End
   Begin VB.Label lblPort 
      Height          =   255
      Index           =   1
      Left            =   1920
      TabIndex        =   12
      Top             =   1920
      Width           =   5415
   End
   Begin VB.Label lblPortDesc 
      Height          =   255
      Index           =   1
      Left            =   1920
      TabIndex        =   11
      Top             =   2280
      Width           =   5415
   End
   Begin VB.Label lblDest 
      Height          =   255
      Index           =   1
      Left            =   1920
      TabIndex        =   10
      Top             =   1560
      Width           =   5415
   End
   Begin VB.Label lblProgram 
      Height          =   255
      Index           =   1
      Left            =   1920
      TabIndex        =   9
      Top             =   1200
      Width           =   5415
   End
   Begin VB.Label lblPortDesc 
      Caption         =   "Trojan Info:"
      Height          =   255
      Index           =   0
      Left            =   720
      TabIndex        =   8
      Top             =   2280
      Width           =   1215
   End
   Begin VB.Label lblPort 
      Caption         =   "Remote Port:"
      Height          =   255
      Index           =   0
      Left            =   720
      TabIndex        =   7
      Top             =   1920
      Width           =   975
   End
   Begin VB.Label lblDest 
      Caption         =   "Destination:"
      Height          =   255
      Index           =   0
      Left            =   720
      TabIndex        =   6
      Top             =   1560
      Width           =   975
   End
   Begin VB.Label lblProgram 
      Caption         =   "Program:"
      Height          =   255
      Index           =   0
      Left            =   720
      TabIndex        =   5
      Top             =   1200
      Width           =   735
   End
   Begin VB.Label lblDESC 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   $"frmAlert.frx":000C
      ForeColor       =   &H80000008&
      Height          =   975
      Left            =   720
      TabIndex        =   0
      Top             =   120
      Width           =   5535
   End
   Begin VB.Image imgPic 
      Height          =   480
      Left            =   120
      Picture         =   "frmAlert.frx":00D0
      Top             =   120
      Width           =   480
   End
End
Attribute VB_Name = "frmAlert"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public bWhatToDo                    As Boolean
Public bRemember                    As Boolean
Private Sub cmdNo_Click()
    Hide
    bWhatToDo = False
    bRemember = (chkRemember.Value = vbChecked)
End Sub
Private Sub cmdYes_Click()
    Hide
    bWhatToDo = True
    bRemember = (chkRemember.Value = vbChecked)
End Sub

