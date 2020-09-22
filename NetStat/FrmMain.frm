VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Begin VB.Form FrmMain 
   Appearance      =   0  'Flat
   Caption         =   "Fire Gate"
   ClientHeight    =   6585
   ClientLeft      =   9450
   ClientTop       =   8310
   ClientWidth     =   9435
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "FrmMain.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6585
   ScaleWidth      =   9435
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer tmrRemoveStatus 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   120
      Top             =   6480
   End
   Begin VB.Timer tmrIcon 
      Interval        =   5000
      Left            =   7080
      Top             =   6840
   End
   Begin VB.PictureBox Pic32 
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Left            =   840
      ScaleHeight     =   32
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   32
      TabIndex        =   13
      Top             =   6840
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.PictureBox Pic16 
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   1440
      ScaleHeight     =   240
      ScaleWidth      =   240
      TabIndex        =   12
      Top             =   6960
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Timer tmrConnections 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   600
      Top             =   6480
   End
   Begin MSComctlLib.ImageList ImgL32 
      Left            =   2160
      Top             =   6840
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      MaskColor       =   -2147483633
      _Version        =   393216
   End
   Begin MSComctlLib.ImageList ImgL16 
      Left            =   1560
      Top             =   6840
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      MaskColor       =   -2147483633
      _Version        =   393216
   End
   Begin MSComctlLib.ImageList ilTray 
      Left            =   3960
      Top             =   6480
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   13
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMain.frx":030A
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMain.frx":0BE4
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMain.frx":14BE
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMain.frx":1D98
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMain.frx":2672
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMain.frx":2F4C
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMain.frx":3826
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMain.frx":4100
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMain.frx":49DA
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMain.frx":4CF4
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMain.frx":500E
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMain.frx":5B58
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMain.frx":66A2
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ilPacket 
      Left            =   4680
      Top             =   6480
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   2
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMain.frx":67FC
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMain.frx":7346
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Frame frmOptions 
      BorderStyle     =   0  'None
      Height          =   5655
      Left            =   240
      TabIndex        =   1
      Top             =   600
      Visible         =   0   'False
      Width           =   8895
      Begin VB.Frame Frame6 
         Appearance      =   0  'Flat
         Caption         =   "General Options"
         ForeColor       =   &H80000008&
         Height          =   1335
         Left            =   120
         TabIndex        =   78
         Top             =   3720
         Width           =   8775
         Begin VB.CheckBox chkIcons 
            Appearance      =   0  'Flat
            Caption         =   "Show Process Icons ?"
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   240
            TabIndex        =   80
            Top             =   480
            Width           =   4575
         End
         Begin VB.CheckBox chkName 
            Appearance      =   0  'Flat
            Caption         =   "Don't resolve hostnames (Increases Performance)"
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   240
            TabIndex        =   79
            Top             =   240
            Width           =   4695
         End
      End
      Begin VB.Frame Frame4 
         Appearance      =   0  'Flat
         Caption         =   "IP Tools"
         ForeColor       =   &H80000008&
         Height          =   1455
         Left            =   6120
         TabIndex        =   27
         Top             =   240
         Width           =   2775
         Begin VB.TextBox txtIPTools 
            Appearance      =   0  'Flat
            Height          =   285
            Index           =   1
            Left            =   120
            TabIndex        =   29
            Top             =   1080
            Width           =   2535
         End
         Begin VB.TextBox txtIPTools 
            Appearance      =   0  'Flat
            Height          =   285
            Index           =   0
            Left            =   120
            TabIndex        =   28
            Top             =   240
            Width           =   2535
         End
         Begin prjNetStat.lvButtons_H cmdHost 
            Height          =   375
            Left            =   120
            TabIndex        =   30
            Top             =   600
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   661
            Caption         =   "Host"
            CapAlign        =   2
            BackStyle       =   3
            Shape           =   2
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Verdana"
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
         Begin prjNetStat.lvButtons_H cmdIP 
            Height          =   375
            Left            =   840
            TabIndex        =   31
            Top             =   600
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   661
            Caption         =   "IP"
            CapAlign        =   2
            BackStyle       =   3
            Shape           =   3
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Verdana"
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
         Begin prjNetStat.lvButtons_H cmdClear 
            Height          =   375
            Left            =   1560
            TabIndex        =   32
            Top             =   600
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   661
            Caption         =   "Clear"
            CapAlign        =   2
            BackStyle       =   3
            Shape           =   1
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Verdana"
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
      End
      Begin VB.Frame Frame2 
         Appearance      =   0  'Flat
         Caption         =   "IP Blocker"
         ForeColor       =   &H80000008&
         Height          =   3375
         Left            =   2640
         TabIndex        =   19
         Top             =   240
         Width           =   3375
         Begin VB.TextBox txtIP 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   120
            TabIndex        =   20
            ToolTipText     =   "Enter an ip or hostmask to deny connections to and from, followed by clicking on Add."
            Top             =   2400
            Width           =   1815
         End
         Begin prjNetStat.lvButtons_H cmdClearIP 
            Height          =   375
            Left            =   2040
            TabIndex        =   21
            ToolTipText     =   "Clear the list of ips to block connections to and from."
            Top             =   2880
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   661
            Caption         =   "Clear"
            CapAlign        =   2
            BackStyle       =   3
            Shape           =   1
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Verdana"
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
         Begin prjNetStat.lvButtons_H cmdAddIP 
            Height          =   375
            Left            =   120
            TabIndex        =   22
            ToolTipText     =   "Add's the ip or hostmask entered in the input box above to the list of denied hosts."
            Top             =   2880
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   661
            Caption         =   "Add"
            CapAlign        =   2
            BackStyle       =   3
            Shape           =   2
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Verdana"
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
         Begin prjNetStat.lvButtons_H cmdDelIP 
            Height          =   375
            Left            =   960
            TabIndex        =   23
            ToolTipText     =   "Delete and item from the IP Blocker list."
            Top             =   2880
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   661
            Caption         =   "Del"
            CapAlign        =   2
            BackStyle       =   3
            Shape           =   3
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Verdana"
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
         Begin prjNetStat.lvButtons_H chkInOut 
            Height          =   375
            Index           =   2
            Left            =   2040
            TabIndex        =   24
            Top             =   2400
            Width           =   735
            _ExtentX        =   1296
            _ExtentY        =   661
            Caption         =   "In"
            CapAlign        =   2
            BackStyle       =   3
            Shape           =   2
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            cGradient       =   0
            Mode            =   1
            Value           =   0   'False
            cBack           =   -2147483633
         End
         Begin prjNetStat.lvButtons_H chkInOut 
            Height          =   375
            Index           =   3
            Left            =   2400
            TabIndex        =   25
            Top             =   2400
            Width           =   735
            _ExtentX        =   1296
            _ExtentY        =   661
            Caption         =   "Out"
            CapAlign        =   2
            BackStyle       =   3
            Shape           =   1
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            cGradient       =   0
            Mode            =   1
            Value           =   0   'False
            cBack           =   -2147483633
         End
         Begin MSComctlLib.ListView lstIPs 
            Height          =   2055
            Left            =   120
            TabIndex        =   26
            ToolTipText     =   "List of current TCP Connections."
            Top             =   240
            Width           =   3135
            _ExtentX        =   5530
            _ExtentY        =   3625
            View            =   3
            LabelEdit       =   1
            Sorted          =   -1  'True
            LabelWrap       =   -1  'True
            HideSelection   =   -1  'True
            FullRowSelect   =   -1  'True
            PictureAlignment=   4
            _Version        =   393217
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            BorderStyle     =   1
            Appearance      =   0
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            NumItems        =   0
         End
      End
      Begin prjNetStat.lvButtons_H cmdDelPort 
         Height          =   375
         Left            =   720
         TabIndex        =   6
         ToolTipText     =   "Delete a Port Number from the list of blocked ports."
         Top             =   3120
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   661
         Caption         =   "Del"
         CapAlign        =   2
         BackStyle       =   3
         Shape           =   3
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
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
      Begin VB.Frame Frame3 
         Appearance      =   0  'Flat
         Caption         =   "Password"
         ForeColor       =   &H80000008&
         Height          =   1455
         Left            =   6120
         TabIndex        =   8
         Top             =   2160
         Width           =   2775
         Begin VB.TextBox txtPassword 
            Appearance      =   0  'Flat
            Height          =   285
            IMEMode         =   3  'DISABLE
            Left            =   120
            PasswordChar    =   "*"
            TabIndex        =   10
            Top             =   360
            Width           =   2535
         End
         Begin prjNetStat.lvButtons_H cmdPassword 
            Height          =   375
            Left            =   840
            TabIndex        =   9
            Top             =   840
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   661
            Caption         =   "Enabled"
            CapAlign        =   2
            BackStyle       =   3
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            cGradient       =   0
            Mode            =   1
            Value           =   0   'False
            cBack           =   -2147483633
         End
      End
      Begin VB.Frame Frame1 
         Appearance      =   0  'Flat
         Caption         =   "Port Blocker"
         ForeColor       =   &H80000008&
         Height          =   3375
         Left            =   120
         TabIndex        =   2
         Top             =   240
         Width           =   2415
         Begin prjNetStat.lvButtons_H chkInOut 
            Height          =   375
            Index           =   0
            Left            =   1080
            TabIndex        =   16
            Top             =   2400
            Width           =   735
            _ExtentX        =   1296
            _ExtentY        =   661
            Caption         =   "In"
            CapAlign        =   2
            BackStyle       =   3
            Shape           =   2
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            cGradient       =   0
            Mode            =   1
            Value           =   0   'False
            cBack           =   -2147483633
         End
         Begin prjNetStat.lvButtons_H cmdClearPort 
            Height          =   375
            Left            =   1320
            TabIndex        =   5
            ToolTipText     =   "Clear the list of Port Numbers and stop blocking traffic on all ports."
            Top             =   2880
            Width           =   975
            _ExtentX        =   1720
            _ExtentY        =   661
            Caption         =   "Clear"
            CapAlign        =   2
            BackStyle       =   3
            Shape           =   1
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Verdana"
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
         Begin prjNetStat.lvButtons_H cmdAddPort 
            Height          =   375
            Left            =   120
            TabIndex        =   4
            ToolTipText     =   "Add's the Port Number in the box above to the list of ports to block traffic on."
            Top             =   2880
            Width           =   855
            _ExtentX        =   1508
            _ExtentY        =   661
            Caption         =   "Add"
            CapAlign        =   2
            BackStyle       =   3
            Shape           =   2
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Verdana"
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
         Begin VB.TextBox txtPort 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   120
            TabIndex        =   7
            ToolTipText     =   "Port Number you wish to block. Type it here followed by clicking on Add."
            Top             =   2440
            Width           =   855
         End
         Begin prjNetStat.lvButtons_H chkInOut 
            Height          =   375
            Index           =   1
            Left            =   1440
            TabIndex        =   17
            Top             =   2400
            Width           =   735
            _ExtentX        =   1296
            _ExtentY        =   661
            Caption         =   "Out"
            CapAlign        =   2
            BackStyle       =   3
            Shape           =   1
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            cGradient       =   0
            Mode            =   1
            Value           =   0   'False
            cBack           =   -2147483633
         End
         Begin MSComctlLib.ListView lstPorts 
            Height          =   2055
            Left            =   120
            TabIndex        =   18
            ToolTipText     =   "List of current TCP Connections."
            Top             =   240
            Width           =   2175
            _ExtentX        =   3836
            _ExtentY        =   3625
            View            =   3
            LabelEdit       =   1
            Sorted          =   -1  'True
            LabelWrap       =   -1  'True
            HideSelection   =   -1  'True
            FullRowSelect   =   -1  'True
            PictureAlignment=   4
            _Version        =   393217
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            BorderStyle     =   1
            Appearance      =   0
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            NumItems        =   0
         End
      End
   End
   Begin MSComctlLib.TabStrip tsType 
      Height          =   6135
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   9255
      _ExtentX        =   16325
      _ExtentY        =   10821
      TabWidthStyle   =   2
      TabFixedWidth   =   2734
      TabMinWidth     =   1025
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   5
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "TCP"
            Object.ToolTipText     =   "List of open TCP Connections."
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "UDP"
            Object.ToolTipText     =   "List of open UDP Connections."
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab3 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Program Control"
            Object.ToolTipText     =   "List of all programs currently being monitered and their status."
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab4 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Options"
            Object.ToolTipText     =   "Program Options"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab5 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Packet Sniffer"
            Object.ToolTipText     =   "Packet Sniffer Logs (If Active)"
            ImageVarType    =   2
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSComctlLib.ListView lstPrograms 
      Height          =   5415
      Left            =   240
      TabIndex        =   15
      Top             =   720
      Visible         =   0   'False
      Width           =   8895
      _ExtentX        =   15690
      _ExtentY        =   9551
      View            =   3
      LabelEdit       =   1
      Sorted          =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      PictureAlignment=   4
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   0
   End
   Begin MSComctlLib.ListView lstUDPConnections 
      Height          =   5415
      Left            =   240
      TabIndex        =   14
      ToolTipText     =   "List of current UDP Connections."
      Top             =   720
      Visible         =   0   'False
      Width           =   8895
      _ExtentX        =   15690
      _ExtentY        =   9551
      View            =   3
      LabelEdit       =   1
      Sorted          =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FlatScrollBar   =   -1  'True
      FullRowSelect   =   -1  'True
      PictureAlignment=   4
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   0
   End
   Begin MSComctlLib.ListView lstTCPConnections 
      Height          =   5415
      Left            =   240
      TabIndex        =   3
      ToolTipText     =   "List of current TCP Connections."
      Top             =   720
      Width           =   8895
      _ExtentX        =   15690
      _ExtentY        =   9551
      View            =   3
      LabelEdit       =   1
      Sorted          =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      PictureAlignment=   4
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   0
   End
   Begin VB.Frame frmConnectionInfo 
      Caption         =   "Connection Information"
      Height          =   2055
      Left            =   360
      TabIndex        =   33
      Top             =   4200
      Visible         =   0   'False
      Width           =   8775
      Begin VB.Frame frmHex 
         BorderStyle     =   0  'None
         Height          =   1700
         Left            =   120
         TabIndex        =   63
         Top             =   240
         Visible         =   0   'False
         Width           =   8535
         Begin VB.PictureBox picOffSet1 
            Appearance      =   0  'Flat
            AutoRedraw      =   -1  'True
            BeginProperty Font 
               Name            =   "Terminal"
               Size            =   9
               Charset         =   255
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   1335
            Left            =   0
            ScaleHeight     =   87
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   26
            TabIndex        =   67
            TabStop         =   0   'False
            Top             =   360
            Width           =   420
         End
         Begin VB.PictureBox picOffset2 
            Appearance      =   0  'Flat
            AutoRedraw      =   -1  'True
            BeginProperty Font 
               Name            =   "Terminal"
               Size            =   9
               Charset         =   255
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   315
            Left            =   405
            ScaleHeight     =   19
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   385
            TabIndex        =   66
            TabStop         =   0   'False
            Top             =   0
            Width           =   5805
         End
         Begin VB.PictureBox picHexDisp 
            Appearance      =   0  'Flat
            AutoRedraw      =   -1  'True
            BackColor       =   &H80000005&
            BeginProperty Font 
               Name            =   "Terminal"
               Size            =   9
               Charset         =   255
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   1335
            Left            =   405
            ScaleHeight     =   87
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   385
            TabIndex        =   65
            TabStop         =   0   'False
            Top             =   360
            Width           =   5805
         End
         Begin VB.PictureBox picChrDisp 
            Appearance      =   0  'Flat
            AutoRedraw      =   -1  'True
            BackColor       =   &H80000005&
            BeginProperty Font 
               Name            =   "Terminal"
               Size            =   9
               Charset         =   255
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   1335
            Left            =   6285
            ScaleHeight     =   87
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   135
            TabIndex        =   64
            Top             =   360
            Width           =   2055
         End
         Begin VB.Label lblFileSize 
            Alignment       =   2  'Center
            BorderStyle     =   1  'Fixed Single
            Caption         =   "lblFileSize"
            ForeColor       =   &H00000080&
            Height          =   315
            Left            =   1440
            TabIndex        =   76
            ToolTipText     =   "File size"
            Top             =   0
            Visible         =   0   'False
            Width           =   1305
         End
         Begin VB.Label lblFileSpec 
            BorderStyle     =   1  'Fixed Single
            Caption         =   "lblFileSpec"
            ForeColor       =   &H00000080&
            Height          =   285
            Left            =   600
            TabIndex        =   75
            Top             =   0
            Visible         =   0   'False
            Width           =   4725
         End
         Begin VB.Label lblCloseHex 
            Caption         =   "X"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   8400
            TabIndex        =   74
            Top             =   0
            Width           =   135
         End
         Begin VB.Label lblHexDownAll 
            Caption         =   "»"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   8400
            TabIndex        =   73
            Top             =   1440
            Width           =   135
         End
         Begin VB.Label lblHexUpAll 
            Caption         =   "«"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   8370
            TabIndex        =   72
            Top             =   360
            Width           =   135
         End
         Begin VB.Label lblHexUp 
            Caption         =   "<"
            Height          =   255
            Left            =   8370
            TabIndex        =   71
            Top             =   600
            Width           =   135
         End
         Begin VB.Label lblHexDown 
            Caption         =   ">"
            Height          =   255
            Left            =   8370
            TabIndex        =   70
            Top             =   1200
            Width           =   135
         End
         Begin VB.Label lblBinary 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   315
            Left            =   7005
            TabIndex        =   69
            ToolTipText     =   "Binary"
            Top             =   0
            Width           =   1305
         End
         Begin VB.Label lblAscii 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   315
            Left            =   6285
            TabIndex        =   68
            ToolTipText     =   "ASCII"
            Top             =   0
            Width           =   705
         End
      End
      Begin VB.Frame frmIPInfo 
         Caption         =   "IP Information"
         Height          =   615
         Left            =   120
         TabIndex        =   58
         Top             =   1320
         Width           =   8415
         Begin VB.Label lblIP_Registrant 
            Caption         =   "Working..."
            Height          =   255
            Left            =   2400
            TabIndex        =   62
            Top             =   240
            Width           =   5775
         End
         Begin VB.Label lblIP_Country 
            Caption         =   "..."
            Height          =   255
            Left            =   960
            TabIndex        =   61
            Top             =   240
            Width           =   375
         End
         Begin VB.Label Label7 
            Caption         =   "Registrant:"
            Height          =   255
            Left            =   1440
            TabIndex        =   60
            Top             =   240
            Width           =   975
         End
         Begin VB.Label Label6 
            Caption         =   "Country:"
            Height          =   255
            Left            =   120
            TabIndex        =   59
            Top             =   240
            Width           =   855
         End
      End
      Begin VB.Label lblCloseInfo 
         Caption         =   "X"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   8520
         TabIndex        =   50
         Top             =   120
         Width           =   135
      End
      Begin VB.Label Label11 
         Alignment       =   1  'Right Justify
         Caption         =   "Server Access:"
         Height          =   255
         Left            =   3360
         TabIndex        =   34
         Top             =   1080
         Width           =   1335
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Connection:"
         Height          =   255
         Left            =   480
         TabIndex        =   49
         Top             =   360
         Width           =   1095
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         Caption         =   "Direction:"
         Height          =   255
         Left            =   480
         TabIndex        =   48
         Top             =   600
         Width           =   1095
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         Caption         =   "Program:"
         Height          =   255
         Left            =   3840
         TabIndex        =   47
         Top             =   360
         Width           =   855
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         Caption         =   "Port Description:"
         Height          =   255
         Left            =   120
         TabIndex        =   46
         Top             =   840
         Width           =   1455
      End
      Begin VB.Label lblCn_Connection 
         Height          =   255
         Left            =   1680
         TabIndex        =   45
         Top             =   360
         Width           =   1935
      End
      Begin VB.Label lblCn_Direction 
         Height          =   255
         Left            =   1680
         TabIndex        =   44
         Top             =   600
         Width           =   1935
      End
      Begin VB.Label lblCn_PortDesc 
         Height          =   255
         Left            =   1680
         TabIndex        =   43
         Top             =   840
         Width           =   1935
      End
      Begin VB.Label Label8 
         Alignment       =   1  'Right Justify
         Caption         =   "Time Initiated:"
         Height          =   255
         Left            =   120
         TabIndex        =   42
         Top             =   1080
         Width           =   1455
      End
      Begin VB.Label Label9 
         Alignment       =   1  'Right Justify
         Caption         =   "Location:"
         Height          =   255
         Left            =   3840
         TabIndex        =   41
         Top             =   600
         Width           =   855
      End
      Begin VB.Label Label10 
         Alignment       =   1  'Right Justify
         Caption         =   "Access:"
         Height          =   255
         Left            =   3840
         TabIndex        =   40
         Top             =   840
         Width           =   855
      End
      Begin VB.Label lblCn_Time 
         Height          =   255
         Left            =   1680
         TabIndex        =   39
         Top             =   1080
         Width           =   1935
      End
      Begin VB.Label lblCn_Location 
         Height          =   255
         Left            =   4800
         TabIndex        =   37
         Top             =   600
         Width           =   3855
      End
      Begin VB.Label lblCn_Access 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   4800
         TabIndex        =   36
         Top             =   840
         Width           =   3855
      End
      Begin VB.Label lblCn_Server 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   4800
         TabIndex        =   35
         Top             =   1080
         Width           =   3855
      End
      Begin VB.Label lblCn_Program 
         Height          =   255
         Left            =   4800
         TabIndex        =   38
         Top             =   360
         Width           =   3855
      End
   End
   Begin VB.Frame Frame5 
      BorderStyle     =   0  'None
      Height          =   5655
      Left            =   240
      TabIndex        =   51
      Top             =   600
      Visible         =   0   'False
      Width           =   8895
      Begin VB.TextBox txtIPLog 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   0
         TabIndex        =   54
         Top             =   260
         Width           =   2535
      End
      Begin VB.ComboBox cmbInterface 
         Height          =   315
         Left            =   2760
         Style           =   2  'Dropdown List
         TabIndex        =   53
         Top             =   260
         Width           =   2550
      End
      Begin prjNetStat.lvButtons_H cmdStartLog 
         Height          =   375
         Left            =   5640
         TabIndex        =   55
         Top             =   240
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   661
         Caption         =   "Start"
         CapAlign        =   2
         BackStyle       =   3
         Shape           =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cGradient       =   0
         Mode            =   2
         Value           =   0   'False
         cBack           =   -2147483633
      End
      Begin prjNetStat.lvButtons_H cmdStopLog 
         Height          =   375
         Left            =   6600
         TabIndex        =   56
         Top             =   240
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   661
         Caption         =   "Stop"
         CapAlign        =   2
         BackStyle       =   3
         Shape           =   3
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cGradient       =   0
         Mode            =   2
         Value           =   -1  'True
         cBack           =   -2147483633
      End
      Begin prjNetStat.lvButtons_H cmdClearPacket 
         Height          =   375
         Left            =   7560
         TabIndex        =   57
         Top             =   240
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   661
         Caption         =   "Clear"
         CapAlign        =   2
         BackStyle       =   3
         Shape           =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
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
      Begin MSComctlLib.ListView lstPacket 
         Height          =   4335
         Left            =   0
         TabIndex        =   52
         ToolTipText     =   "List of current TCP Connections."
         Top             =   720
         Width           =   8895
         _ExtentX        =   15690
         _ExtentY        =   7646
         View            =   3
         Arrange         =   1
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         PictureAlignment=   2
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   0
      End
      Begin VB.Label Label5 
         Caption         =   "Filter:"
         Height          =   255
         Left            =   0
         TabIndex        =   77
         Top             =   0
         Width           =   1215
      End
   End
   Begin VB.Label lblStatus 
      Alignment       =   1  'Right Justify
      Height          =   255
      Left            =   120
      TabIndex        =   11
      Top             =   0
      Width           =   9135
   End
   Begin VB.Menu mnFile 
      Caption         =   "&File"
      Begin VB.Menu mnRegClean 
         Caption         =   "|Clean the registry of all data belonging to Fire Gate|Clean &Registry"
      End
      Begin VB.Menu mnSep0 
         Caption         =   "-"
         Index           =   0
      End
      Begin VB.Menu mnCloseAll 
         Caption         =   "|Close Fire Gate|&Close"
      End
   End
   Begin VB.Menu mnView 
      Caption         =   "&View"
      Begin VB.Menu mnViewTCP 
         Caption         =   "|View all current TCP Connections|&TCP Connections"
      End
      Begin VB.Menu mnViewUDP 
         Caption         =   "|View all current UDP Connections|&UDP Connections"
      End
      Begin VB.Menu mnViewPrograms 
         Caption         =   "|View all Programs and their rules|&Program List"
      End
      Begin VB.Menu mnViewOptions 
         Caption         =   "|View Program Options|&Options"
      End
      Begin VB.Menu mnViewPackets 
         Caption         =   "|View packets in real time as they are sent/received|&Packet Sniffer"
      End
   End
   Begin VB.Menu mnPopup 
      Caption         =   "Popup"
      Visible         =   0   'False
      Begin VB.Menu mnPopShow 
         Caption         =   "&Hide"
      End
      Begin VB.Menu mnChar 
         Caption         =   "(-)Exit"
      End
      Begin VB.Menu mnPopExit 
         Caption         =   "&Exit"
      End
   End
   Begin VB.Menu mnRightClick 
      Caption         =   "Right Click Menu"
      Visible         =   0   'False
      Begin VB.Menu mnInformation 
         Caption         =   "|View Information about this Connection|View &Information"
      End
      Begin VB.Menu mnCopyInfo 
         Caption         =   "|Copy various information about this connection to the clipboard.|&Copy"
         Begin VB.Menu mnCopyRemoteIP 
            Caption         =   "Remote IP"
         End
         Begin VB.Menu mnCopyLocalIP 
            Caption         =   "Local IP"
         End
         Begin VB.Menu mnSep3 
            Caption         =   "(-)Program"
         End
         Begin VB.Menu mnCopyProgName 
            Caption         =   "Program Name"
         End
         Begin VB.Menu mnCopyProgLoc 
            Caption         =   "Program Location"
         End
      End
      Begin VB.Menu mnPacketSniff 
         Caption         =   "|Sniff packets on this connection.|&Packet Sniff"
         Begin VB.Menu mnLocalIP 
            Caption         =   "|Sniff Connection using Local IP as filter.|&Local IP"
         End
         Begin VB.Menu mnSniffRemote 
            Caption         =   "|Sniff Connection using Remote IP as filter.|&Remote IP"
         End
      End
      Begin VB.Menu mnClose 
         Caption         =   "|Close|Close.."
         Begin VB.Menu mnCloseProgram 
            Caption         =   "|Close the owner process.|Close &Program"
         End
         Begin VB.Menu mnCloseConnection 
            Caption         =   "|Close this current connection.|C&lose Connection"
         End
         Begin VB.Menu mnSep1 
            Caption         =   "-"
         End
         Begin VB.Menu mnCloseBoth 
            Caption         =   "|Close both the program and connection.|&Both"
         End
      End
   End
   Begin VB.Menu mnHex 
      Caption         =   "mnHex"
      Visible         =   0   'False
      Begin VB.Menu mnHexView 
         Caption         =   "|View data in a new window.|&View"
         Begin VB.Menu mnViewHex 
            Caption         =   "|View hex data in a new window.|H&ex"
         End
         Begin VB.Menu mnViewText 
            Caption         =   "|View text data in a new window.|Te&xt"
         End
      End
      Begin VB.Menu mnCopy 
         Caption         =   "|Copy data to the clipboard.|&Copy"
         Begin VB.Menu mnCopyText 
            Caption         =   "|Copy text to the clipboard.|&Text"
         End
         Begin VB.Menu mnCopyHex 
            Caption         =   " |Copy hex to the clipboard.|&Hex"
         End
      End
   End
   Begin VB.Menu mnHelpMenu 
      Caption         =   "Help"
      Begin VB.Menu mnAbout 
         Caption         =   "|Information about Fire Gate and its author.|&About"
      End
      Begin VB.Menu mnHelpSep 
         Caption         =   "(-)Help"
      End
      Begin VB.Menu mnHelp 
         Caption         =   "|Program Help|&Help"
      End
   End
End
Attribute VB_Name = "FrmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Threading = http://www.Planet-Source-Code.com/vb/scripts/ShowCode.asp?txtCodeId=26900&lngWId=1
Private WithEvents FormAlert        As MThreadVB.Thread
Attribute FormAlert.VB_VarHelpID = -1
Private WithEvents pThread          As MThreadVB.Thread
Attribute pThread.VB_VarHelpID = -1
Private ProtocolBuilder             As clsProtocolInterface
Private WithEvents TCPDriver        As clsTCPProtocol
Attribute TCPDriver.VB_VarHelpID = -1
Private WithEvents UDPDriver        As clsUDPProtocol
Attribute UDPDriver.VB_VarHelpID = -1
Private WithEvents IPInformation    As clsIPInfo
Attribute IPInformation.VB_VarHelpID = -1
Private WithEvents HelpObj          As clsHelpCallBack
Attribute HelpObj.VB_VarHelpID = -1
Private sLoggingConnection          As String
Private lstMenu                     As ListView
Private iCurrentIcon                As Integer
Sub ShowBarForm(ModalFlag As Variant)
    FormAlert.ObjectInThreadContext.ShowFormNow 1
End Sub
Sub ShowFormNow(ModalFlag As Long)
    FrmBar.Show
End Sub
Sub ShowThreadForm(ModalFlag As Variant)
    pThread.ObjectInThreadContext.ShowThreadNow 1
End Sub
Sub ShowThreadNow(ModalFlag As Long)
    FrmThread.Show
End Sub
Private Sub chkInOut_Click(Index As Integer)
    Select Case Index
        Case Is < 2
            If chkInOut(0).Value = False And chkInOut(1).Value = False Then chkInOut(0).Value = True
        Case Else
            If chkInOut(2).Value = False And chkInOut(3).Value = False Then chkInOut(2).Value = True
    End Select
End Sub
Private Sub cmdAddIP_Click()
    On Error GoTo ErrClear
    Dim Item                        As ListItem
    Dim sTmp                        As String
    Dim X                           As Integer
    sTmp = txtIP.Text
    If Len(sTmp) > 0 Then
        If chkInOut(2).Value = False And chkInOut(3).Value = False Then
            Call MsgBox("You must select to block this ip on a combination of Incoming, Outgoing or Both !", vbExclamation, App.Title)
            GoTo ErrClear
        End If
        Set Item = lstIPs.ListItems.Add(, "IP-" & sTmp, sTmp)
        If chkInOut(2).Value = True And chkInOut(3).Value = True Then
            Item.ListSubItems.Add(, , "Both").Tag = "0"
        ElseIf chkInOut(2).Value = True And chkInOut(3).Value = False Then
            Item.ListSubItems.Add(, , "In").Tag = "1"
        ElseIf chkInOut(2).Value = False And chkInOut(3).Value = True Then
            Item.ListSubItems.Add(, , "Out").Tag = "2"
        End If
        SaveIPs
ErrClear:
        If Err.Number = MSComctlLib.ErrorConstants.ccNonUniqueKey Then Call MsgBox("You are already blocking traffic to/from " & sTmp & " !", vbExclamation, App.Title)
        Err.Clear
        Set Item = Nothing
    End If
End Sub
Private Sub cmdAddPort_Click()
    On Error GoTo ErrClear
    Dim Item                        As ListItem
    Dim sTmp                        As String
    Dim X                           As Integer
    sTmp = txtPort.Text
    If Len(sTmp) > 0 Then
        If chkInOut(0).Value = False And chkInOut(1).Value = False Then
            Call MsgBox("You must select to block this port on a combination of Incoming, Outgoing or Both !", vbExclamation, App.Title)
            GoTo ErrClear
        End If
        Set Item = lstPorts.ListItems.Add(, "Port-" & txtPort.Text, txtPort.Text)
        If chkInOut(0).Value = True And chkInOut(1).Value = True Then
            Item.ListSubItems.Add(, , "Both").Tag = "0"
        ElseIf chkInOut(0).Value = True And chkInOut(1).Value = False Then
            Item.ListSubItems.Add(, , "In").Tag = "1"
        ElseIf chkInOut(0).Value = False And chkInOut(1).Value = True Then
            Item.ListSubItems.Add(, , "Out").Tag = "2"
        End If
        SavePorts
ErrClear:
        If Err.Number = MSComctlLib.ErrorConstants.ccNonUniqueKey Then Call MsgBox("You are already blocking traffic on port " & sTmp & " !", vbExclamation, App.Title)
        Err.Clear
        Set Item = Nothing
    End If
End Sub
Private Sub cmdClear_Click()
    txtIPTools(1).Text = ""
    txtIPTools(0).Text = ""
End Sub
Private Sub cmdClearIP_Click()
    Dim lRet                        As VbMsgBoxResult
    If lstIPs.ListItems.Count > 0 Then
        lRet = MsgBox("Are you sure you want to stop blocking traffic to and from all the ips/hostmasks in the list ?", vbQuestion Or vbYesNo, App.Title)
        If lRet = vbYes Then
            lstIPs.ListItems.Clear
            SaveIPs
        End If
    End If
End Sub
Private Sub cmdClearPacket_Click()
    lstPacket.ListItems.Clear
End Sub
Private Sub cmdClearPort_Click()
    Dim lRet                        As VbMsgBoxResult
    If lstPorts.ListItems.Count > 0 Then
        lRet = MsgBox("Are you sure you want to stop blocking traffic on ALL ports ?", vbQuestion Or vbYesNo, App.Title)
        If lRet = vbYes Then
            lstPorts.ListItems.Clear
            SavePorts
        End If
    End If
End Sub
Private Sub cmdDelIP_Click()
    On Error GoTo ErrClear
    Dim Item                        As ListItem
    Dim lRet                        As VbMsgBoxResult
    Dim iRet                        As Integer
    Set Item = lstIPs.SelectedItem
    iRet = Len(Item.Text)
    lRet = MsgBox("Do you really want to stop blocking traffic to/from " & Item.Text & "?", vbYesNo Or vbQuestion, App.Title)
    If lRet = vbYes Then
        lstIPs.ListItems.Remove Item.Index
        SaveIPs
    End If
ErrClear:
    If Err.Number = 91 Then Call MsgBox("You must first select an ip or hostmask to remove before clicking the remove button !", vbExclamation, App.Title)
    Err.Clear
    Set Item = Nothing
End Sub
Private Sub cmdDelPort_Click()
    On Error GoTo ErrClear
    Dim Item                        As ListItem
    Dim lRet                        As VbMsgBoxResult
    Dim iRet                        As Integer
    Set Item = lstPorts.SelectedItem
    iRet = Len(Item.Text)
    lRet = MsgBox("Do you really want to stop blocking traffic on port " & Item.Text & "?", vbYesNo Or vbQuestion, App.Title)
    If lRet = vbYes Then
        lstPorts.ListItems.Remove Item.Index
        SavePorts
    End If
ErrClear:
    If Err.Number = 91 Then Call MsgBox("You must first select a port to remove before clicking the remove button !", vbExclamation, App.Title)
    Err.Clear
    Set Item = Nothing
End Sub
Private Sub cmdHost_Click()
    txtIPTools(1).Text = ModIptoHost.NameByAddr(txtIPTools(0).Text)
End Sub
Private Sub cmdIP_Click()
    txtIPTools(1).Text = ModIptoHost.NameToAddr(txtIPTools(0).Text)
End Sub
Private Sub cmdPassword_Click()
    Debug.Print GetShortPath("C:\Program Files\Steam\ClientRegistry.blob")
End Sub

Private Sub cmdStartLog_Click()
    sLoggingConnection = Left$(cmbInterface.Text, InStr(1, cmbInterface, " "))
    If ProtocolBuilder.CreateRawSocket(sLoggingConnection, 7000, Me.hwnd) = 0 Then
        txtIPLog.Enabled = False
        cmdStopLog.Value = True
        cmdStartLog.Value = False
    End If
End Sub
Private Sub cmdStopLog_Click()
    ProtocolBuilder.CloseRawSocket
End Sub

Private Sub Form_Load()
    Dim Str()                       As String
    Dim X                           As Integer
    
    Set FormAlert = New Thread
    Set pThread = New Thread
    If FormAlert.IsThreadRunning = False Then FormAlert.CreateWin32Thread Me, "ShowBarForm", 0
    If pThread.IsThreadRunning = False Then pThread.CreateWin32Thread Me, "ShowThreadForm", 0

    Set ProtocolBuilder = New clsProtocolInterface
    Set TCPDriver = New clsTCPProtocol
    Set UDPDriver = New clsUDPProtocol
    Set g_aProgramDescriptions = New Dictionary
    Set g_DBCon = MakeADOConnection
    Set g_rsTrojan = New ADODB.Recordset
    Set g_rsPorts = New ADODB.Recordset
    Proc_Startup
    ProtocolBuilder.AddinProtocol TCPDriver, "TCP", IPPROTO_TCP
    ProtocolBuilder.AddinProtocol UDPDriver, "UDP", IPPROTO_UDP
    
    Str = Split(EnumNetworkInterfaces(), ";")
    For X = 0 To UBound(Str)
        If Str(X) <> "127.0.0.1" Then
            cmbInterface.AddItem Str(X) & " [" & GetHostNameByAddr(inet_addr(Str(X))) & "]"
        End If
    Next
    If cmbInterface.ListCount > 0 Then cmbInterface.ListIndex = 0
    
    Set HelpObj = New clsHelpCallBack
    Call ModCoolMenu.Install(Me.hwnd, HelpObj)
    ModCoolMenu.ForeColor (Me.hwnd)
    Call ModCoolMenu.FullSelect(Me.hwnd, True)
    
    txtIP.ToolTipText = txtIP.ToolTipText & " (Note: * wildcard is supported)"
    
    LoadSettings
    LoadPrograms
    
    Pic16.Width = (SMALL_ICON) * Screen.TwipsPerPixelX 'Set the Temp Picture Box properties.
    Pic16.Height = (SMALL_ICON) * Screen.TwipsPerPixelY 'Set the Temp Picture Box properties.
    Pic32.Width = LARGE_ICON * Screen.TwipsPerPixelX 'Set the Temp Picture Box properties.
    Pic32.Height = LARGE_ICON * Screen.TwipsPerPixelY 'Set the Temp Picture Box properties.
    g_sShell32Path = FixPath(SpecialFolder(WinSystem)) & "shell32.dll" 'Get the shell32.dll location into memory for later use.
    IP_Initialize
    With lstTCPConnections
        If g_bXPTable = True Then .ColumnHeaders.Add , , "Program", 2500
        .ColumnHeaders.Add , , "Local Address", 2500
        .ColumnHeaders.Add , , "Local Port", 1100
        .ColumnHeaders.Add , , "Remote Address", 2500
        .ColumnHeaders.Add , , "Remote Port", 1300
        .ColumnHeaders.Add , , "State", 1000
        If g_bXPTable = True Then
            .ColumnHeaders.Add , , "User", 1000
            .ColumnHeaders.Add , , "Process ID", 1200
        End If
        .ZOrder
    End With
    With lstUDPConnections
        If g_bXPTable = True Then .ColumnHeaders.Add , , "Program", 2500
        .ColumnHeaders.Add , , "Local Address", 2500
        .ColumnHeaders.Add , , "Local Port", 1100
        If g_bXPTable = True Then
            .ColumnHeaders.Add , , "User", 1000
            .ColumnHeaders.Add , , "Process ID", 1200
        End If
    End With
    With lstPrograms
        .Icons = ilTray
        .SmallIcons = ilTray
        .ColumnHeaders.Add , , "Active", 700
        .ColumnHeaders.Add , , "Program", 3500
        .ColumnHeaders.Add , , "Access", 800
        .ColumnHeaders.Add , , "Server", 800
    End With
    With lstPorts
        .ColumnHeaders.Add , , "Port", 850
        .ColumnHeaders.Add , , "Direction", 1300
    End With
    With lstIPs
        .ColumnHeaders.Add , , "IP/Host", 1700
        .ColumnHeaders.Add , , "Direction"
    End With
    With lstPacket
        .SmallIcons = ilPacket
        .ColumnHeaders.Add , , "Source", 2500
        .ColumnHeaders.Add , , "Destination", 2500
        .ColumnHeaders.Add , , "Time", 1200
        .ColumnHeaders.Add , , "Ver", 800
        .ColumnHeaders.Add , , "Data", 2500
    End With
    MakeNumberOnly txtPort
    LoadPorts
    LoadIPs
    Form_Resize
    
    With nidProgramData
        .cbSize = Len(nidProgramData)
        .hwnd = FrmMenu.hwnd
        .uId = vbNull
        .uFlags = NIF_ICON Or NIF_TIP Or NIF_MESSAGE
        .uCallbackMessage = WM_MOUSEMOVE
        .hIcon = FrmMain.ilTray.ListImages(10).ExtractIcon
        .szTip = "Fire Gate" & vbNullChar
    End With
    Shell_NotifyIcon NIM_ADD, nidProgramData
    
    FrmThread.tmrThread.Enabled = True
    FrmThread.Visible = False
End Sub
Private Sub Form_Resize()
    On Error GoTo ErrClear
    If WindowState = vbMinimized Then
        If Visible = True Then
            mnPopShow.Caption = "|Show The Main Window|&Show"
            Visible = False
        End If
        Exit Sub
    Else
        If Visible = False Then Visible = True
    End If
    If Width < 9555 Then Width = 9555: Exit Sub
    If Height < 7215 Then Height = 7215: Exit Sub
    Dim iTmp                        As Integer
    tsType.Height = Height - tsType.Top - 1000
    tsType.Width = Width - 360
    iTmp = tsType.Height - 600
    Frame5.Height = iTmp
    If Not frmConnectionInfo.Visible Then
        lstTCPConnections.Height = iTmp
        lstUDPConnections.Height = iTmp
        lstPrograms.Height = iTmp
        lstPacket.Height = iTmp - 720
    Else
        lstTCPConnections.Height = iTmp - 2200
        lstUDPConnections.Height = iTmp - 2200
        lstPrograms.Height = iTmp - 2200
        lstPacket.Height = iTmp - 2920
    End If
    frmOptions.Height = iTmp
    
    iTmp = tsType.Width - 240
    lstTCPConnections.Width = iTmp
    lstUDPConnections.Width = iTmp
    lstPrograms.Width = iTmp
    frmOptions.Width = iTmp
    Frame5.Width = iTmp
    lstPacket.Width = iTmp - 80
    'lblStatus.Top = Me.Height - lblStatus.Height
    
    iTmp = frmConnectionInfo.Width
    frmConnectionInfo.Width = Width - 360 - 420
    frmConnectionInfo.Top = Height - 3200
    frmIPInfo.Width = frmConnectionInfo.Width - 275
    lblIP_Registrant.Width = frmIPInfo.Width - 2500
    iTmp = iTmp - frmConnectionInfo.Width
    
    frmHex.Width = frmConnectionInfo.Width - 240
    'iTmp = frmHex.Width - 135
    MoveThing lblCloseHex, iTmp
    MoveThing lblHexDown, iTmp
    MoveThing lblHexDownAll, iTmp
    MoveThing lblHexUp, iTmp
    MoveThing lblHexUpAll, iTmp
    MoveThing picChrDisp, iTmp
    MoveThing picHexDisp, iTmp
    MoveThing picOffSet1, iTmp
    MoveThing picOffset2, iTmp
    MoveThing lblAscii, iTmp
    MoveThing lblBinary, iTmp
    'lblCloseHex.Left = iTmp
    'lblHexDown.Left = iTmp
    'lblHexDownAll.Left = iTmp
    'lblHexUp.Left = iTmp
    'lblHexUpAll.Left = iTmp
    
    MoveThing Label3, iTmp
    MoveThing lblCloseInfo, iTmp, , True
    
    MoveThing Label9, iTmp
    MoveThing Label10, iTmp
    MoveThing Label11, iTmp
    
    MoveThing lblCn_Connection, iTmp, True
    MoveThing lblCn_Direction, iTmp, True
    MoveThing lblCn_PortDesc, iTmp, True
    MoveThing lblCn_Time, iTmp, True
    
    MoveThing lblCn_Access, iTmp
    MoveThing lblCn_Access, iTmp, True
    MoveThing lblCn_Location, iTmp
    MoveThing lblCn_Location, iTmp, True
    MoveThing lblCn_Program, iTmp
    MoveThing lblCn_Program, iTmp, True
    MoveThing lblCn_Server, iTmp
    MoveThing lblCn_Server, iTmp, True
    Exit Sub
ErrClear:
    lblStatus.Caption = "Status: Form_Resize(Error : #" & Err.Number & " " & Err.Description & ")"
    Err.Clear
End Sub
Private Sub MoveThing(lbl As Control, iMove As Integer, Optional bWidth As Boolean = False, Optional bNotDivide As Boolean = False)
    If bWidth Then
        lbl.Width = lbl.Width - (iMove / 2)
    Else
        lbl.Left = lbl.Left - IIf(Not bNotDivide, (iMove / 2), iMove)
    End If
End Sub
Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
    Dim lRet                        As VbMsgBoxResult
    lRet = MsgBox("Are you sure you want to close " & App.Title & " ?" & vbCrLf & "This will disable your personal firewall and your computer will be susceptible to attack !", vbQuestion Or vbYesNo, App.Title)
    If lRet = vbNo Then
        Cancel = 1
        Exit Sub
    End If
    
    g_bStopAll = True
    Do While g_bWorking = True
        FrmThread.DoThread
        DoEvents
    Loop
    SaveSettings
    ModCoolMenu.Uninstall Me.hwnd
    
    tmrConnections.Enabled = False
    tmrIcon.Enabled = False
    ProtocolBuilder.CloseRawSocket
    Set ProtocolBuilder = Nothing
    Set TCPDriver = Nothing
    Set UDPDriver = Nothing
    Set IPInformation = Nothing
    
    If g_rsPorts.State = adStateOpen Then g_rsPorts.Close
    Set g_rsPorts = Nothing
    If g_rsTrojan.State = adStateOpen Then g_rsTrojan.Close
    Set g_rsTrojan = Nothing
    If g_DBCon.State = adStateOpen Then g_DBCon.Close
    Set g_DBCon = Nothing
    
    Erase g_aPrograms
    Erase g_TcpConnections
    Erase g_UdpConnections
    Unload FrmBar
    Unload FrmThread
    g_aProgramDescriptions.RemoveAll
    Set g_aProgramDescriptions = Nothing
    Proc_Cleanup
    Set pThread = Nothing
    Set FormAlert = Nothing
    Set lstMenu = Nothing
    WSACleanup
    Unload FrmMenu
    Unload Me
    If InIDE = False Then End
End Sub
Private Sub HelpObj_MenuHelp(ByVal MenuText As String, ByVal MenuHelp As String, ByVal Enabled As Boolean)
    If Len(MenuHelp) > 0 Then
        lblStatus.Caption = MenuHelp
        tmrRemoveStatus.Enabled = False
        tmrRemoveStatus.Enabled = True
    End If
End Sub
Private Sub IPInformation_Finished(sCompany As String, sCountry As String, sIP As String)
    lblIP_Country.Caption = sCountry
    lblIP_Registrant.Caption = sCompany
    Set IPInformation = Nothing
End Sub
Private Sub lblCloseHex_Click()
    frmConnectionInfo.Visible = False
    frmConnectionInfo.Caption = "Connection Information"
    lblCloseInfo.Visible = True
    frmHex.Visible = False
    picChrDisp.Picture = LoadPicture()
    picHexDisp.Picture = LoadPicture()
    picOffSet1.Picture = LoadPicture()
    picOffset2.Picture = LoadPicture()
    lblAscii.Caption = ""
    lblBinary.Caption = ""
    Form_Resize
End Sub
Private Sub lblCloseInfo_Click()
    frmConnectionInfo.Visible = False
    Form_Resize
End Sub
Private Sub lblHexDown_Click()
    HexMoveDown
End Sub
Private Sub lblHexDownAll_Click()
    HexMovePageDown
End Sub
Private Sub lblHexUp_Click()
    HexMoveUp
End Sub
Private Sub lblHexUpAll_Click()
    HexMovePageUp
End Sub
Private Sub lstPacket_DblClick()
    On Error GoTo ErrClear
    Dim Item                        As ListItem
    Set Item = lstPacket.SelectedItem
    If Len(Item.Tag) = 0 Then Exit Sub
    lblCloseInfo.Visible = False
    frmHex.Visible = True
    frmConnectionInfo.Visible = True
    frmConnectionInfo.Caption = "Packet Data"
    frmConnectionInfo.ZOrder
    Form_Resize
    ModHexDump.HexStartup
    LoadData Item.Tag
    Set Item = Nothing
ErrClear:
End Sub
Private Sub lstTCPConnections_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    lstTCPConnections.Sorted = True 'Sort it.
    lstTCPConnections.SortKey = ColumnHeader.Index - 1
    lstTCPConnections.SortOrder = IIf(lstTCPConnections.SortOrder = lvwAscending, lvwDescending, lvwAscending)  'SortOrder = Not SortOrder
End Sub
Private Sub lstTCPConnections_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    'Checked = Remove
    If Button = 2 Then
        Set lstMenu = lstTCPConnections
        PopupMenu mnRightClick
    End If
End Sub
Private Sub lstUDPConnections_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    lstUDPConnections.Sorted = True 'Sort it.
    lstUDPConnections.SortKey = ColumnHeader.Index - 1
    lstUDPConnections.SortOrder = IIf(lstUDPConnections.SortOrder = lvwAscending, lvwDescending, lvwAscending)  'SortOrder = Not SortOrder    Set lstMenu = lstUDPConnections
End Sub
Private Sub lstUDPConnections_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then
        Set lstMenu = lstUDPConnections
        PopupMenu mnRightClick
    End If
End Sub
Private Sub mnAbout_Click()
    Hide
    frmAboutScreen.Show vbModal, Me
    Show
End Sub
Private Sub mnCloseAll_Click()
    Unload Me
End Sub
Private Sub mnCloseBoth_Click()
    On Error GoTo ErrClear
    Dim Item                        As ListItem
    Dim X                           As Integer
    Dim lProcID                     As Long
    Set Item = lstMenu.SelectedItem
    If lstMenu Is lstTCPConnections Then
        lProcID = CLng(Item.ListSubItems(7).Text)
        For X = 0 To C_UBound(g_TcpConnections)
            If g_TcpConnections(X).ProcInfo.lProcID = lProcID And g_TcpConnections(X).lRemotePort = Item.ListSubItems(4).Text And g_TcpConnections(X).sRemoteAddr = Item.ListSubItems(3).Tag And g_TcpConnections(X).sLocalAddr = Item.ListSubItems(1).Text And g_TcpConnections(X).lLocalPort = Item.ListSubItems(2).Text Then
                CloseConnection g_TcpConnections(X).Row
                KillProcess g_TcpConnections(X).ProcInfo.lProcID
                Exit For
            End If
        Next
    Else
        lProcID = CLng(Item.ListSubItems(4).Text)
        For X = 0 To C_UBound(g_UdpConnections)
            If g_UdpConnections(X).ProcInfo.lProcID = lProcID And g_UdpConnections(X).sLocalAddr = Item.ListSubItems(1).Text And g_UdpConnections(X).lLocalPort = Item.ListSubItems(2).Text Then
                CloseConnection g_UdpConnections(X).Row
                KillProcess g_UdpConnections(X).ProcInfo.lProcID
                Exit For
            End If
        Next
    End If
ErrClear:
    Err.Clear
End Sub
Private Sub mnCloseConnection_Click()
    On Error GoTo ErrClear
    Dim Item                        As ListItem
    Dim X                           As Integer
    Dim lProcID                     As Long
    Set Item = lstMenu.SelectedItem
    If lstMenu Is lstTCPConnections Then
        lProcID = CLng(Item.ListSubItems(7).Text)
        For X = 0 To C_UBound(g_TcpConnections)
            If g_TcpConnections(X).ProcInfo.lProcID = lProcID And g_TcpConnections(X).lRemotePort = Item.ListSubItems(4).Text And g_TcpConnections(X).sRemoteAddr = Item.ListSubItems(3).Tag And g_TcpConnections(X).sLocalAddr = Item.ListSubItems(1).Text And g_TcpConnections(X).lLocalPort = Item.ListSubItems(2).Text Then
                CloseConnection g_TcpConnections(X).Row
                Exit For
            End If
        Next
    Else
        lProcID = CLng(Item.ListSubItems(4).Text)
        For X = 0 To C_UBound(g_UdpConnections)
            If g_UdpConnections(X).ProcInfo.lProcID = lProcID And g_UdpConnections(X).sLocalAddr = Item.ListSubItems(1).Text And g_UdpConnections(X).lLocalPort = Item.ListSubItems(2).Text Then
                CloseConnection g_UdpConnections(X).Row
                Exit For
            End If
        Next
    End If
ErrClear:
    Err.Clear
End Sub
Private Sub mnCloseProgram_Click()
    On Error GoTo ErrClear
    Dim Item                        As ListItem
    Dim X                           As Integer
    Dim lProcID                     As Long
    Set Item = lstMenu.SelectedItem
    If lstMenu Is lstTCPConnections Then
        lProcID = CLng(Item.ListSubItems(7).Text)
        For X = 0 To C_UBound(g_TcpConnections)
            If g_TcpConnections(X).ProcInfo.lProcID = lProcID And g_TcpConnections(X).lRemotePort = Item.ListSubItems(4).Text And g_TcpConnections(X).sRemoteAddr = Item.ListSubItems(3).Tag And g_TcpConnections(X).sLocalAddr = Item.ListSubItems(1).Text And g_TcpConnections(X).lLocalPort = Item.ListSubItems(2).Text Then
                KillProcess g_TcpConnections(X).ProcInfo.lProcID
                Exit For
            End If
        Next
    Else
        lProcID = CLng(Item.ListSubItems(4).Text)
        For X = 0 To C_UBound(g_UdpConnections)
            If g_UdpConnections(X).ProcInfo.lProcID = lProcID And g_UdpConnections(X).sLocalAddr = Item.ListSubItems(1).Text And g_UdpConnections(X).lLocalPort = Item.ListSubItems(2).Text Then
                KillProcess g_UdpConnections(X).ProcInfo.lProcID
                Exit For
            End If
        Next
    End If
ErrClear:
    Err.Clear
End Sub
Private Sub mnCopyHex_Click()
    CopyHexData (False)
End Sub

Private Sub mnCopyLocalIP_Click()
    Dim Item                        As ListItem
    On Error GoTo ErrClear
    Set Item = lstMenu.SelectedItem
    If Len(Item.Text) > 0 Then
        Clipboard.Clear
        Clipboard.SetText Item.ListSubItems(1).Text
    End If
    Set Item = Nothing
    Exit Sub
ErrClear:
    Err.Clear
    Set Item = Nothing
End Sub
Private Sub mnCopyProgLoc_Click()
    Dim Item                        As ListItem
    On Error GoTo ErrClear
    Set Item = lstMenu.SelectedItem
    If Len(Item.Text) > 0 Then
        Clipboard.Clear
        Clipboard.SetText Item.ListSubItems(2).Tag
    End If
    Set Item = Nothing
    Exit Sub
ErrClear:
    Err.Clear
    Set Item = Nothing
End Sub
Private Sub mnCopyProgName_Click()
    Dim Item                        As ListItem
    On Error GoTo ErrClear
    Set Item = lstMenu.SelectedItem
    If Len(Item.Text) > 0 Then
        Clipboard.Clear
        Clipboard.SetText Item.Text
    End If
    Set Item = Nothing
    Exit Sub
ErrClear:
    Err.Clear
    Set Item = Nothing
End Sub
Private Sub mnCopyRemoteIP_Click()
    Dim Item                        As ListItem
    On Error GoTo ErrClear
    Set Item = lstMenu.SelectedItem
    If Len(Item.Text) > 0 Then
        Clipboard.Clear
        Clipboard.SetText Item.ListSubItems(3).Text
    End If
    Set Item = Nothing
    Exit Sub
ErrClear:
    Err.Clear
    Set Item = Nothing
End Sub
Private Sub mnCopyText_Click()
    CopyHexData (True)
End Sub
Private Sub mnInformation_Click()
    Dim Item                        As ListItem
    Dim X                           As Integer
    Dim i                           As Integer
    Dim lProcID                     As Long
    Dim sTmp                        As String
    On Error GoTo ErrClear
    Set Item = lstMenu.SelectedItem
    lblIP_Country.Caption = "..."
    lblIP_Registrant.Caption = "Working..."
    Set IPInformation = New clsIPInfo
    If lstMenu Is lstTCPConnections Then
        lProcID = CLng(Item.ListSubItems(7).Text)
        For X = 0 To C_UBound(g_TcpConnections)
            If g_TcpConnections(X).ProcInfo.lProcID = lProcID And g_TcpConnections(X).lRemotePort = Item.ListSubItems(4).Text And g_TcpConnections(X).sRemoteAddr = Item.ListSubItems(3).Tag And g_TcpConnections(X).sLocalAddr = Item.ListSubItems(1).Text And g_TcpConnections(X).lLocalPort = Item.ListSubItems(2).Text Then
                IPInformation.GetDetails g_TcpConnections(X).sRemoteAddr
                lblCn_Connection.Caption = IIf(g_TcpConnections(X).Direction = Incoming, Item.ListSubItems(3).Text & ":" & Item.ListSubItems(4).Text & " -> " & Item.ListSubItems(1).Text & ":" & Item.ListSubItems(2).Text, Item.ListSubItems(1).Text & ":" & Item.ListSubItems(2).Text & " -> " & Item.ListSubItems(3).Text & ":" & Item.ListSubItems(4).Text)
                lblCn_Program.Caption = g_aProgramDescriptions(LCase(g_TcpConnections(X).ProcInfo.sPath))
                lblCn_Location.Caption = g_TcpConnections(X).ProcInfo.sPath  'g_aProcessTable(lProcID)
                lblCn_Direction.Caption = IIf(g_TcpConnections(X).Direction = Incoming, "Incoming", "Outgoing")
                sTmp = PortDetails(CStr(g_TcpConnections(X).lRemotePort), IIf(g_TcpConnections(X).Direction = Incoming, True, False), TCP)
                If Len(sTmp) > 0 Then
                    lblCn_PortDesc.Caption = sTmp & " (" & CStr(g_TcpConnections(X).lRemotePort) & ")"
                Else
                    lblCn_PortDesc.Caption = "[No Description Available]" & " (" & CStr(g_TcpConnections(X).lRemotePort) & ")"
                End If
                lblCn_Time.Caption = Item.ListSubItems(1).Tag
                For i = 0 To T_UBound(g_aPrograms)
                    If LCase(g_TcpConnections(X).ProcInfo.sPath) = g_aPrograms(i).sLocation Then
                        Select Case g_aPrograms(i).iAccess
                            Case 0
                                lblCn_Access.Caption = "Ask"
                                lblCn_Access.ForeColor = vbMagenta
                            Case 1
                                lblCn_Access.Caption = "Deny"
                                lblCn_Access.ForeColor = vbRed
                            Case 2
                                lblCn_Access.Caption = "Allow"
                                lblCn_Access.ForeColor = vbGreen
                        End Select
                        Select Case g_aPrograms(i).iServer
                            Case 0
                                lblCn_Server.Caption = "Ask"
                                lblCn_Server.ForeColor = vbMagenta
                            Case 1
                                lblCn_Server.Caption = "Deny"
                                lblCn_Server.ForeColor = vbRed
                            Case 2
                                lblCn_Server.Caption = "Allow"
                                lblCn_Server.ForeColor = vbGreen
                        End Select
                        Exit For
                    End If
                Next
                Exit For
            End If
        Next
    Else
        lProcID = CLng(Item.ListSubItems(4).Text)
        For X = 0 To C_UBound(g_UdpConnections)
            If g_UdpConnections(X).ProcInfo.lProcID = lProcID And g_UdpConnections(X).sLocalAddr = Item.ListSubItems(1).Text And g_UdpConnections(X).lLocalPort = Item.ListSubItems(2).Text Then
                IPInformation.GetDetails g_UdpConnections(X).sRemoteAddr
                lblCn_Connection.Caption = Item.ListSubItems(1).Text & ":" & Item.ListSubItems(2).Text
                lblCn_Program.Caption = g_aProgramDescriptions(LCase(g_UdpConnections(X).ProcInfo.sPath))
                lblCn_Location.Caption = g_UdpConnections(X).ProcInfo.sPath
                lblCn_Direction.Caption = "UDP"
                sTmp = PortDetails(CStr(g_UdpConnections(X).lRemotePort), False, UDP)
                If Len(sTmp) > 0 Then
                    lblCn_PortDesc.Caption = sTmp & " (" & CStr(g_UdpConnections(X).lLocalPort) & ")"
                Else
                    lblCn_PortDesc.Caption = "[No Description Available]" & " (" & CStr(g_UdpConnections(X).lLocalPort) & ")"
                End If
                lblCn_Time.Caption = Item.ListSubItems(1).Tag
                For i = 0 To T_UBound(g_aPrograms)
                    If LCase(g_UdpConnections(X).ProcInfo.sPath) = g_aPrograms(i).sLocation Then
                        Select Case g_aPrograms(i).iAccess
                            Case 0
                                lblCn_Access.Caption = "Ask"
                                lblCn_Access.ForeColor = vbMagenta
                            Case 1
                                lblCn_Access.Caption = "Deny"
                                lblCn_Access.ForeColor = vbRed
                            Case 2
                                lblCn_Access.Caption = "Allow"
                                lblCn_Access.ForeColor = vbGreen
                        End Select
                        Select Case g_aPrograms(i).iServer
                            Case 0
                                lblCn_Server.Caption = "Ask"
                                lblCn_Server.ForeColor = vbMagenta
                            Case 1
                                lblCn_Server.Caption = "Deny"
                                lblCn_Server.ForeColor = vbRed
                            Case 2
                                lblCn_Server.Caption = "Allow"
                                lblCn_Server.ForeColor = vbGreen
                        End Select
                        Exit For
                    End If
                Next
                Exit For
            End If
        Next
    End If
    frmConnectionInfo.Caption = "Connection Information"
    lblCloseInfo.Visible = True
    frmHex.Visible = False
    frmConnectionInfo.ZOrder
    frmConnectionInfo.Visible = True
    Form_Resize
ErrClear:
    Err.Clear
End Sub
Private Sub mnLocalIP_Click()
    Dim Item                        As ListItem
    On Error GoTo ErrClear
    Set Item = lstMenu.SelectedItem
    If Len(Item.Text) > 0 Then
        txtIPLog.Text = Item.ListSubItems(1).Text
        tsType.Tabs(5).Selected = True
    End If
    Exit Sub
ErrClear:
    Err.Clear
End Sub
Private Sub mnPopExit_Click()
    Unload Me
End Sub
Private Sub mnPopShow_Click()
    Dim sTmp                        As String
    sTmp = Right$(mnPopShow.Caption, 4)
    If StrComp(sTmp, "Hide", vbBinaryCompare) = 0 Then
        Me.WindowState = vbMinimized
        Visible = False
    Else
        mnPopShow.Caption = "|Hide The Main Window|&Hide"
        WindowState = vbNormal
        Visible = True
        SetForegroundWindow hwnd
    End If
End Sub
Private Sub mnRegClean_Click()
    If REGDeleteSetting(vHKEY_LOCAL_MACHINE, "Software\EliteProdigy\Fire Gate", vbNullString) = True Then
        lblStatus.Caption = "Regitsry Cleaning succeeded !"
    Else
        lblStatus.Caption = "Regitsry Cleaning Failed !"
    End If
End Sub

Private Sub mnSniffRemote_Click()
    Dim Item                        As ListItem
    On Error GoTo ErrClear
    Set Item = lstMenu.SelectedItem
    If Len(Item.Text) > 0 Then
        txtIPLog.Text = Item.ListSubItems(3).Text
        tsType.Tabs(5).Selected = True
    End If
    Exit Sub
ErrClear:
    Err.Clear
End Sub
Private Sub mnViewOptions_Click()
    tsType.Tabs(4).Selected = True
End Sub
Private Sub mnViewPackets_Click()
    tsType.Tabs(5).Selected = True
End Sub
Private Sub mnViewPrograms_Click()
    tsType.Tabs(3).Selected = True
End Sub
Private Sub mnViewTCP_Click()
    tsType.Tabs(1).Selected = True
End Sub
Private Sub mnViewUDP_Click()
    tsType.Tabs(2).Selected = True
End Sub
Private Sub picChrDisp_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = vbLeftButton Then
        Call ModHexDump.picChrDisp_MouseUp(Button, Shift, X, Y)
    Else
        PopupMenu mnHex
    End If
End Sub
Private Sub picHexDisp_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = vbLeftButton Then
        Call ModHexDump.picHexDisp_MouseUp(Button, Shift, X, Y)
    Else
        PopupMenu mnHex
    End If
End Sub
Private Sub TCPDriver_RecievedPacket(IPHeader As clsIPHeader, TCPProtocol As clsTCPProtocol, Data As String, bData() As Byte)
    Dim Item                        As ListItem
    Dim sData                       As String
    Dim sTmp(1)                     As String
    If Len(txtIPLog.Text) > 0 Then
        If chkName.Value = vbUnchecked Then
            sTmp(0) = NameByAddr(IPHeader.DestIP)
            sTmp(1) = NameByAddr(IPHeader.SourceIP)
        End If
        If MatchSpec(IPHeader.DestIP, txtIPLog.Text) = True Or MatchSpec(IPHeader.SourceIP, txtIPLog.Text) = True Then GoTo OkContinue
        If chkName.Value = vbUnchecked Then
            If Len(sTmp(0)) > 0 Then
                If MatchSpec(sTmp(0), txtIPLog.Text) = True Then
                    GoTo OkContinue
                End If
            End If
            If Len(sTmp(1)) > 0 Then
                If MatchSpec(sTmp(1), txtIPLog.Text) = True Then
                    GoTo OkContinue
                End If
            End If
        End If
        Set Item = Nothing
        Exit Sub
    End If
OkContinue:
    If IPHeader.DestIP = sLoggingConnection Then
        Set Item = lstPacket.ListItems.Add(, , IPHeader.DestIP & ":" & CStr(TCPProtocol.DestPort), , 1)
        Item.ListSubItems.Add , , IPHeader.SourceIP & ":" & CStr(TCPProtocol.SourcePort)
    Else
        Set Item = lstPacket.ListItems.Add(, , IPHeader.SourceIP & ":" & CStr(TCPProtocol.SourcePort), , 2)
        Item.ListSubItems.Add , , IPHeader.DestIP & ":" & CStr(TCPProtocol.SourcePort)
    End If
    Item.ListSubItems.Add , , Time$
    Item.ListSubItems.Add , , "IPv" & CStr(IPHeader.Version)
    Item.ListSubItems.Add , , Replace(IIf(Len(Data) <= 35, Data, Left$(Data, 35) & "..."), Chr(0), ".")
    Item.Tag = Replace(Data, Chr(0), ".")
    Set Item = Nothing
End Sub
Private Sub tmrRemoveStatus_Timer()
    lblStatus.Caption = ""
    tmrRemoveStatus.Enabled = False
End Sub
Private Sub UDPDriver_RecievedPacket(IPHeader As clsIPHeader, UDPProtocol As clsUDPProtocol, Data As String)
    Dim Item                        As ListItem
    Dim sData                       As String
    Dim sTmp(1)                     As String
    If Len(txtIPLog.Text) > 0 Then
        If chkName.Value = vbUnchecked Then
            sTmp(0) = NameByAddr(IPHeader.DestIP)
            sTmp(1) = NameByAddr(IPHeader.SourceIP)
        End If
        If MatchSpec(IPHeader.DestIP, txtIPLog.Text) = True Or MatchSpec(IPHeader.SourceIP, txtIPLog.Text) = True Then GoTo OkContinue
        If chkName.Value = vbUnchecked Then
            If Len(sTmp(0)) > 0 Then If MatchSpec(sTmp(0), txtIPLog.Text) = True Then GoTo OkContinue:
            If Len(sTmp(1)) > 0 Then If MatchSpec(sTmp(1), txtIPLog.Text) = True Then GoTo OkContinue:
        End If
        Set Item = Nothing
        Exit Sub
    End If
OkContinue:
    If IPHeader.DestIP = sLoggingConnection Then
        Set Item = lstPacket.ListItems.Add(, , IPHeader.DestIP & ":" & CStr(UDPProtocol.DestPort), , 1)
        Item.ListSubItems.Add , , IPHeader.SourceIP & ":" & CStr(UDPProtocol.SourcePort)
    Else
        Set Item = lstPacket.ListItems.Add(, , IPHeader.SourceIP & ":" & CStr(UDPProtocol.SourcePort), , 2)
        Item.ListSubItems.Add , , IPHeader.DestIP & ":" & CStr(UDPProtocol.SourcePort)
    End If
    Item.ListSubItems.Add , , Time$
    Item.ListSubItems.Add , , "IPv" & CStr(IPHeader.Version)
    Item.ListSubItems.Add , , Replace(IIf(Len(Data) <= 35, Data, Left$(Data, 35) & "..."), Chr(0), ".")
    Item.Tag = Replace(Data, Chr(0), ".")
    Set Item = Nothing
End Sub
Private Sub tmrIcon_Timer()
    If Not iCurrentIcon = 10 Then
        nidProgramData.hIcon = ilTray.ListImages(10).ExtractIcon
        Shell_NotifyIcon NIM_MODIFY, nidProgramData
        iCurrentIcon = 10
    End If
End Sub
Private Sub tsType_Click()
    Static iLastTicked              As Integer
    If iLastTicked = 0 Then iLastTicked = 1
    Select Case True
        Case Is = tsType.Tabs(1).Selected
            If iLastTicked = 1 Then Exit Sub
            frmConnectionInfo.Visible = False
            If frmHex.Visible = True Then
                frmConnectionInfo.Caption = "Connection Information"
                frmHex.Visible = False
                lblCloseInfo.Visible = True
            End If
            Form_Resize
            lstTCPConnections.Visible = True
            lstUDPConnections.Visible = False
            lstPrograms.Visible = False
            lstTCPConnections.ZOrder
            Frame5.Visible = False
            iLastTicked = 1
        Case Is = tsType.Tabs(2).Selected
            If iLastTicked = 2 Then Exit Sub
            frmConnectionInfo.Visible = False
            If frmHex.Visible = True Then
                frmConnectionInfo.Caption = "Connection Information"
                lblCloseInfo.Visible = True
                frmHex.Visible = False
            End If
            Form_Resize
            lstTCPConnections.Visible = False
            Frame5.Visible = False
            lstUDPConnections.Visible = True
            lstPrograms.Visible = False
            lstUDPConnections.ZOrder
            iLastTicked = 2
        Case Is = tsType.Tabs(3).Selected
            If iLastTicked = 3 Then Exit Sub
            frmConnectionInfo.Visible = False
            Frame5.Visible = False
            If frmHex.Visible = True Then
                frmConnectionInfo.Caption = "Connection Information"
                lblCloseInfo.Visible = True
                frmHex.Visible = False
            End If
            Form_Resize
            lstTCPConnections.Visible = False
            lstUDPConnections.Visible = False
            lstPrograms.Visible = True
            lstPrograms.ZOrder
            iLastTicked = 3
        Case Is = tsType.Tabs(4).Selected
            If iLastTicked = 4 Then Exit Sub
            Frame5.Visible = False
            frmConnectionInfo.Visible = False
            If frmHex.Visible = True Then
                frmConnectionInfo.Caption = "Connection Information"
                lblCloseInfo.Visible = True
                frmHex.Visible = False
            End If
            Form_Resize
            lstTCPConnections.Visible = False
            lstUDPConnections.Visible = False
            lstPrograms.Visible = False
            frmOptions.Visible = True
            frmOptions.ZOrder
            iLastTicked = 4
        Case Is = tsType.Tabs(5).Selected
            If iLastTicked = 5 Then Exit Sub
            frmConnectionInfo.Visible = False
            If frmHex.Visible = True Then
                frmConnectionInfo.Caption = "Connection Information"
                lblCloseInfo.Visible = True
                frmHex.Visible = False
            End If
            Form_Resize
            lstTCPConnections.Visible = False
            lstUDPConnections.Visible = False
            lstPrograms.Visible = False
            frmOptions.Visible = False
            Frame5.Visible = True
            Frame5.ZOrder
            iLastTicked = 5
    End Select
    If tsType.Tabs(4).Selected = False And frmOptions.Visible = True Then frmOptions.Visible = False
End Sub
Public Function CheckTraffic(iIncoming() As Integer, iOutgoing() As Integer, iBlocked() As Integer, iCount As Integer)
    If iIncoming(1 + iCount) = iIncoming(0 + iCount) And iOutgoing(1 + iCount) = iOutgoing(0 + iCount) And iBlocked(1 + iCount) <> iBlocked(0 + iCount) Then
        'BLOCK ALL
        If iCurrentIcon <> 2 Then
            nidProgramData.hIcon = ilTray.ListImages(2).ExtractIcon
            Shell_NotifyIcon NIM_MODIFY, nidProgramData
            iCurrentIcon = 2
        End If
    End If
    If iIncoming(1 + iCount) <> iIncoming(0 + iCount) And iOutgoing(1 + iCount) <> iOutgoing(0 + iCount) And iBlocked(1 + iCount) = iBlocked(0 + iCount) Then
        'ALLOW ALL
        If iCurrentIcon <> 5 Then
            nidProgramData.hIcon = ilTray.ListImages(5).ExtractIcon
            Shell_NotifyIcon NIM_MODIFY, nidProgramData
            iCurrentIcon = 5
        End If
    End If
    If iIncoming(1 + iCount) = iIncoming(0 + iCount) And iOutgoing(1 + iCount) <> iOutgoing(0 + iCount) And iBlocked(1 + iCount) <> iBlocked(0 + iCount) Then
        'BLOCK OUT
        If iCurrentIcon <> 3 Then
            nidProgramData.hIcon = ilTray.ListImages(3).ExtractIcon
            Shell_NotifyIcon NIM_MODIFY, nidProgramData
            iCurrentIcon = 3
        End If
    End If
    If iIncoming(1 + iCount) <> iIncoming(0 + iCount) And iOutgoing(1 + iCount) = iOutgoing(0 + iCount) And iBlocked(1 + iCount) <> iBlocked(0 + iCount) Then
        'BLOCK IN
        If iCurrentIcon <> 4 Then
            nidProgramData.hIcon = ilTray.ListImages(4).ExtractIcon
            Shell_NotifyIcon NIM_MODIFY, nidProgramData
            iCurrentIcon = 4
        End If
    End If
    If iIncoming(1) = iIncoming(0) And iOutgoing(1) <> iOutgoing(0) And iBlocked(1) = iBlocked(0) Then
        'ALLOW OUT
        If iCurrentIcon <> 6 Then
            nidProgramData.hIcon = ilTray.ListImages(6).ExtractIcon
            Shell_NotifyIcon NIM_MODIFY, nidProgramData
            iCurrentIcon = 6
        End If
    End If
    If iIncoming(1) <> iIncoming(0) And iOutgoing(1) = iOutgoing(0) And iBlocked(1) = iBlocked(0) Then
        'ALLOW IN
        If iCurrentIcon <> 7 Then
            nidProgramData.hIcon = ilTray.ListImages(7).ExtractIcon
            Shell_NotifyIcon NIM_MODIFY, nidProgramData
            iCurrentIcon = 7
        End If
    End If
End Function
