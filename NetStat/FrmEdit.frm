VERSION 5.00
Begin VB.Form FrmEdit 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Edit Program Settings"
   ClientHeight    =   3195
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4680
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.ComboBox cmbAccess 
      Height          =   315
      ItemData        =   "FrmEdit.frx":0000
      Left            =   1320
      List            =   "FrmEdit.frx":000D
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   1320
      Width           =   2295
   End
   Begin VB.ComboBox cmbServer 
      Height          =   315
      ItemData        =   "FrmEdit.frx":0023
      Left            =   1320
      List            =   "FrmEdit.frx":0030
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   1800
      Width           =   2295
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   975
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   4335
   End
End
Attribute VB_Name = "FrmEdit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

