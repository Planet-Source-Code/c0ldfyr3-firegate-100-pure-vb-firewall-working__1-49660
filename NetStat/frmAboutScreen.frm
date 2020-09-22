VERSION 5.00
Begin VB.Form frmAboutScreen 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Program Information"
   ClientHeight    =   2565
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4575
   BeginProperty Font 
      Name            =   "Comic Sans MS"
      Size            =   14.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   -1  'True
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmAboutScreen.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2565
   ScaleWidth      =   4575
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer timText 
      Interval        =   1
      Left            =   0
      Top             =   480
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "&Ok"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1740
      TabIndex        =   1
      Top             =   2160
      Width           =   1095
   End
   Begin VB.PictureBox picText 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   1935
      Left            =   0
      ScaleHeight     =   1935
      ScaleWidth      =   4575
      TabIndex        =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   4575
   End
   Begin VB.Line lnSpacer 
      X1              =   120
      X2              =   4440
      Y1              =   2040
      Y2              =   2040
   End
End
Attribute VB_Name = "frmAboutScreen"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Declare Function GetTickCount Lib "kernel32" () As Long
Private mstrAllText             As String
Private mblnStart               As Boolean
Private Sub cmdOk_Click()
    Unload Me
End Sub
Private Sub Form_Load()
    Icon = FrmMain.Icon
    Call SetText
End Sub
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Unload Me
End Sub
Private Sub RectToPixels(ByRef TheRect As RECT)
    TheRect.Left = TheRect.Left \ Screen.TwipsPerPixelX
    TheRect.Right = TheRect.Right \ Screen.TwipsPerPixelX
    TheRect.Top = TheRect.Top \ Screen.TwipsPerPixelY
    TheRect.Bottom = TheRect.Bottom \ Screen.TwipsPerPixelY
End Sub
Private Sub timText_Timer()
    Const Wait                  As Integer = 50
    Dim udtFont                 As FontStruc
    Dim udtBmp                  As BitmapStruc
    Dim udtMask                 As BitmapStruc
    Dim udtBmpSize              As RECT
    Dim intResult               As Integer
    Dim intTextHeight           As Integer
    Dim lngStartingTick         As Long
    Static udtSurphase          As BitmapStruc
    Static intScroll            As Integer
    lngStartingTick = GetTickCount
    udtBmpSize.Right = picText.ScaleWidth
    udtBmpSize.Bottom = picText.ScaleHeight
    RectToPixels udtBmpSize
    udtMask.Area = udtBmpSize
    udtSurphase.Area = udtBmpSize
    udtBmp.Area = udtBmpSize
    udtFont.Alignment = vbCentreAlign
    udtFont.Name = picText.FontName
    udtFont.Bold = picText.FontBold
    udtFont.Colour = vbWhite
    udtFont.Italic = picText.FontItalic
    udtFont.StrikeThru = picText.FontStrikethru
    udtFont.PointSize = picText.FontSize
    udtFont.Underline = picText.FontUnderline
    intTextHeight = GetTextHeight(picText.hdc) * LineCount(mstrAllText)
    intScroll = intScroll - Screen.TwipsPerPixelY
    If (intScroll < -(intTextHeight * Screen.TwipsPerPixelY)) _
       Or (Not mblnStart) Then
        intScroll = picText.ScaleHeight
        mblnStart = True
    End If
    If udtSurphase.hDcMemory = 0 Then
        Call CreateNewBitmap(udtSurphase.hDcMemory, udtSurphase.hDcBitmap, udtSurphase.hDcPointer, udtSurphase.Area, frmAboutScreen.hdc, picText.ForeColor, InPixels)
        Call Gradient(udtSurphase.hDcMemory, picText.ForeColor, picText.FillColor, 0, (udtSurphase.Area.Bottom - ((intTextHeight / LineCount(mstrAllText)) * 2)), udtSurphase.Area.Right, (intTextHeight / LineCount(mstrAllText) * 2), GradHorizontal, InPixels)
        Call Gradient(udtSurphase.hDcMemory, picText.FillColor, picText.ForeColor, 0, 0, udtSurphase.Area.Right, (intTextHeight / LineCount(mstrAllText)) * 2, GradHorizontal, InPixels)
    End If
    Call CreateNewBitmap(udtMask.hDcMemory, udtMask.hDcBitmap, udtMask.hDcPointer, udtMask.Area, frmAboutScreen.hdc, vbBlack, InPixels)
    Call CreateNewBitmap(udtBmp.hDcMemory, udtBmp.hDcBitmap, udtBmp.hDcPointer, udtBmp.Area, frmAboutScreen.hdc, vbWhite, InPixels)
    Call MakeText(udtMask.hDcMemory, mstrAllText, (intScroll / Screen.TwipsPerPixelY), 0, intTextHeight, udtBmp.Area.Right, udtFont, InPixels)
    intResult = BitBlt(udtBmp.hDcMemory, 0, 0, udtBmp.Area.Right, udtBmp.Area.Bottom, udtSurphase.hDcMemory, 0, 0, SRCCOPY)
    intResult = BitBlt(udtBmp.hDcMemory, 0, 0, udtBmp.Area.Right, udtBmp.Area.Bottom, udtMask.hDcMemory, 0, 0, SRCAND)
    intResult = BitBlt(frmAboutScreen.hdc, 0, 0, udtBmp.Area.Right, udtBmp.Area.Bottom, udtBmp.hDcMemory, 0, 0, SRCCOPY)
    Call DeleteBitmap(udtBmp.hDcMemory, udtBmp.hDcBitmap, udtBmp.hDcPointer)
    Call DeleteBitmap(udtMask.hDcMemory, udtMask.hDcBitmap, udtMask.hDcPointer)
    Call Pause(Wait - (GetTickCount - lngStartingTick))
End Sub
Private Sub SetText()
    mstrAllText = App.ProductName & vbCrLf & _
                  "Version " & App.Major & "." & App.Minor & "." & App.Revision & vbCrLf & _
                  "" & vbCrLf & _
                  "Idea and programming  by" & vbCrLf & _
                  "Jason Croghan aka c0ldfyr3." & vbCrLf & _
                  "" & vbCrLf & _
                  "Copyright 2003" & vbCrLf & _
                  "All rights reserved" & vbCrLf & _
                  "" & vbCrLf & _
                  "Contact:" & vbCrLf & _
                  "E-Mail : c0ldfyr3@EliteProdigy.com" & vbCrLf & _
                  "WWW : http://www.EliteProdigy.com" & vbCrLf & _
                  "Yahoo : c0ldfyr3" & vbCrLf & _
                  "AIM : c0idfyr3" & vbCrLf & _
                  "Hotmail : c0ldfyr3@Hotmail.com" & vbCrLf & _
                  "ICQ : 126177863" & vbCrLf & vbCrLf & _
                  "Credits" & vbCrLf & "========" & vbCrLf & _
                  "ApiByName - M.A. Munim" & vbCrLf & _
                  "FakeBar - Niknak!!" & vbCrLf & _
                  "Hex Viewer - Herman L" & vbCrLf & _
                  "Socket API - Oleg Gdalevich" & vbCrLf & _
                  "Suspend Program - Ananda Raja" & vbCrLf & _
                  "Cool Menu - Olivier Martin" & vbCrLf & _
                  "Packet Sniffer - Coding Genius" & vbCrLf & vbCrLf & _
                  "If I forgot anyone, sorry" & vbCrLf & "just email me and I will" & vbCrLf & "add you to the list."
End Sub
Private Function LineCount(ByVal strText As String) As Integer
    Dim intTemp                 As Integer
    Dim intCounter              As Integer
    Dim intLastPos              As Integer
    intLastPos = 1
    Do
        intTemp = intLastPos
        intLastPos = InStr(intLastPos + Len(vbCrLf), strText, vbCrLf)
        If intTemp <> intLastPos Then
            intCounter = intCounter + 1
        End If
    Loop Until intLastPos = 0
    LineCount = intCounter
End Function
