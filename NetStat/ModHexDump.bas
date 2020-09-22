Attribute VB_Name = "ModHexDump"
'Herman L,http://www.pscode.com/vb/scripts/ShowCode.asp?txtCodeId=11380&lngWId=1
Option Explicit
Private mFileSize                   As Long
Private mSuspend                    As Boolean
Private pageStart                   As Long
Private pageEnd                     As Long
Private StdW1                       As Long
Private StdH1                       As Long
Private StdW2                       As Long
Private StdH2                       As Long
Private ChrW                        As Long
Private arrByte()                   As Byte
Private sString                     As String
Private sHexString                  As String
Private Const CharsInRow = 14
Private Const CharsInCol = 7
Private Const mPageSize = CharsInRow * CharsInCol
Public Sub HexStartup()
    StdW1 = FrmMain.picHexDisp.ScaleWidth / CharsInRow
    StdH1 = FrmMain.picHexDisp.ScaleHeight / CharsInCol
    StdW2 = FrmMain.picChrDisp.ScaleWidth / CharsInRow
    StdH2 = FrmMain.picChrDisp.ScaleHeight / CharsInCol
    ChrW = FrmMain.picHexDisp.TextWidth("X")
End Sub
Private Sub ShowPage(ByVal Hilit As Boolean, Optional ByVal inStart As Long = 0, Optional ByVal inEnd As Long = 0, Optional ByVal inPaint1 As Long, Optional ByVal inPaint2 As Long)
    On Error Resume Next
    If Len(FrmMain.lblFileSpec.Caption) = 0 Then Exit Sub
    Dim strContent                  As String
    Dim offSetPos                   As String
    Dim unDispChar                  As String
    Dim mAscii                      As Integer
    Dim mHex                        As String
    Dim X                           As Integer
    Dim Y                           As Integer
    Dim tmp                         As String
    Dim i                           As Long
    Dim j                           As Long
    Dim k                           As Long
    Dim origX                       As Single
    Dim origY                       As Single

    FrmMain.picHexDisp.Picture = LoadPicture()
    FrmMain.picChrDisp.Picture = LoadPicture()
    FrmMain.picOffSet1.Picture = LoadPicture()
    FrmMain.picOffset2.Picture = LoadPicture()
    
      ' Since we repaint, any values in ASCII & Binary labels are no longer valid
    FrmMain.lblAscii.Caption = ""
    FrmMain.lblBinary.Caption = ""
    
      ' Adjust if required - safety
    If mFileSize <= mPageSize Then
         pageStart = 1
         pageEnd = mFileSize
    Else
         If pageStart < 1 Then
             pageStart = 1
             pageEnd = mPageSize
             If pageEnd > mFileSize Then pageEnd = mFileSize
         End If
         If pageEnd > mFileSize Then
             k = (mFileSize - 1) / mPageSize
             k = NoFraction(k)
             pageStart = k * mPageSize + 1
             pageEnd = pageStart + mPageSize - 1
             If pageEnd > mPageSize Then pageEnd = mPageSize
         End If
    End If
    
      ' Also adjust if required - safety
    If (inStart > 0 And inStart < pageStart) Then inStart = 0
    If (inEnd > 0 And inEnd > pageEnd) Then inEnd = 0
    
      ' Display offset subhead
    FrmMain.picOffset2.CurrentY = 3
    For X = 0 To 15
        tmp = Format$(X, "@@")
        FrmMain.picOffset2.CurrentX = X * StdW1
        FrmMain.picOffset2.Print tmp;
    Next X
    
      ' Restart from top
    FrmMain.picHexDisp.CurrentX = 0
    FrmMain.picChrDisp.CurrentX = 0
    FrmMain.picOffSet1.CurrentY = 0
    FrmMain.picHexDisp.CurrentY = 0
    FrmMain.picChrDisp.CurrentY = 0
    
    unDispChar = "."
    i = pageStart
    Do While i <= pageEnd
        offSetPos = Format$(i, " @@@@")
        FrmMain.picOffSet1.Print offSetPos
        
        For j = 0 To 15
            If (i + j) > pageEnd Or (i + j) > mFileSize Then
                Exit For
            Else
                mAscii = arrByte(i + j)
                   ' For Hex area
                mHex = Hex(mAscii)
                If Len(mHex) < 2 Then
                    mHex = "0" & mHex
                End If
                
                FrmMain.picHexDisp.CurrentX = j * StdW1
                
                If Hilit = True And (inStart > 0 And inEnd > 0) Then
                    If (i + j) >= inStart And (i + j) <= inEnd Then
                        origX = FrmMain.picHexDisp.CurrentX
                        origY = FrmMain.picHexDisp.CurrentY
                        X = j * StdW1 - ChrW * 0.4
                        Y = FrmMain.picHexDisp.CurrentY
                        FrmMain.picHexDisp.ForeColor = inPaint1
                        FrmMain.picHexDisp.Line (X, Y)-(X + ChrW * 2.8, Y + FrmMain.picHexDisp.TextHeight("X")), , BF                                  '"XX" + 0.4*2
                        FrmMain.picHexDisp.ForeColor = inPaint2
                        FrmMain.picHexDisp.CurrentX = origX
                        FrmMain.picHexDisp.CurrentY = origY
                        FrmMain.picHexDisp.Print mHex;
                        FrmMain.picHexDisp.ForeColor = vbBlack
                    Else
                        FrmMain.picHexDisp.Print mHex;
                    End If
                Else
                    FrmMain.picHexDisp.Print mHex;
                End If
                
                X = j * StdW2
                FrmMain.picChrDisp.CurrentX = X
                
                   ' For Chr area
                If Hilit = True And (inStart > 0 And inEnd > 0) Then
                    If (i + j) >= inStart And (i + j) <= inEnd Then
                        origX = FrmMain.picChrDisp.CurrentX
                        origY = FrmMain.picChrDisp.CurrentY
                        
                        Y = FrmMain.picChrDisp.CurrentY
                        FrmMain.picChrDisp.ForeColor = inPaint1
                        FrmMain.picChrDisp.Line (X, Y)-(X + ChrW, Y + FrmMain.picChrDisp.TextHeight("X")), inPaint1, BF                                 ' "X"
                        FrmMain.picChrDisp.ForeColor = inPaint2
                        FrmMain.picChrDisp.CurrentX = origX
                        FrmMain.picChrDisp.CurrentY = origY
                        
                        If mAscii > 31 Then
                            FrmMain.picChrDisp.Print Chr(mAscii);
                        Else
                            FrmMain.picChrDisp.Print unDispChar;
                        End If
                        FrmMain.picChrDisp.ForeColor = vbBlack
                    Else
                        If mAscii > 31 Then
                            FrmMain.picChrDisp.Print Chr(mAscii);
                        Else
                            FrmMain.picChrDisp.Print unDispChar;
                        End If
                    End If
                Else
                    If mAscii > 31 Then
                        FrmMain.picChrDisp.Print Chr(mAscii);
                    Else
                        FrmMain.picChrDisp.Print unDispChar;
                    End If
                End If
            End If
        Next j
        i = i + CharsInRow
        FrmMain.picHexDisp.Print               ' Force picHexDisp change row after earlier ";"
        FrmMain.picHexDisp.CurrentX = 0
        FrmMain.picChrDisp.Print
        FrmMain.picChrDisp.CurrentX = 0        ' Force picChrDisp change row after earlier ";"
    Loop
    For i = 1 To 3
        FrmMain.picHexDisp.Line (StdW1 * 4 * i - ChrW * 0.6, 0)-(StdW1 * 4 * i - ChrW * 0.6, 3), vbBlue, BF
    Next i
    For i = 1 To 3
        FrmMain.picHexDisp.Line (StdW1 * 4 * i - ChrW * 0.6, FrmMain.picHexDisp.ScaleHeight - 4)-(StdW1 * 4 * i - ChrW * 0.6, FrmMain.picHexDisp.ScaleHeight), vbBlue, BF
    Next i
End Sub
Function NoFraction(ByVal inVal As Variant) As Long
    Dim X                           As Integer
    Dim sTmp                        As String
    Dim k                           As Long
    sTmp = CStr(inVal)
    X = InStr(sTmp, ".")
    If X > 0 Then sTmp = Left(sTmp, X - 1)
    k = val(sTmp)
    NoFraction = k
End Function
Public Sub CopyHexData(Optional bString As Boolean = True)
    Clipboard.Clear
    If bString = True Then
        Clipboard.SetText sString
    Else
        Clipboard.SetText sHexString
    End If
End Sub
Public Function LoadData(sData As String)
    On Error GoTo ErrHandler
    Dim X                           As Long
    Dim sTmp                        As String
    Dim aHex()                      As String
    Dim aStr()                      As String
    Dim mHex                        As String
    mFileSize = Len(sData)
    If mFileSize = 0 Then Exit Function
    ReDim arrByte(1 To mFileSize)
    ReDim aHex(Len(sData) - 1)
    ReDim aStr(Len(sData) - 1)
    For X = 1 To Len(sData)
        sTmp = Mid(sData, X, 1)
        arrByte(X) = AscB(sTmp)
        mHex = Hex(arrByte(X))
        If Len(mHex) < 2 Then mHex = "0" & mHex
        aHex(X - 1) = mHex
        If Asc(sTmp) < 32 Then sTmp = "."
        aStr(X - 1) = sTmp
    Next
    sHexString = Join(aHex, " ")
    sString = Join(aStr)
    FrmMain.lblFileSize.Caption = CStr(mFileSize) & " bytes"
    FrmMain.lblFileSpec.Caption = Space(2) & "Some File"
    pageStart = 1
    pageEnd = mPageSize
    ShowPage True
    FrmMain.picHexDisp.SetFocus
    Exit Function
ErrHandler:
    If Err.Number <> 32755 Then
         Screen.MousePointer = vbDefault
         FrmMain.lblFileSize.Caption = ""
         FrmMain.lblFileSpec.Caption = ""
         FrmMain.picHexDisp.Picture = LoadPicture()
         FrmMain.picChrDisp.Picture = LoadPicture()
         FrmMain.picOffSet1.Picture = LoadPicture()
         FrmMain.picOffset2.Picture = LoadPicture()
    End If
End Function
Private Function GetByteIndex(ByVal X As Single, ByVal Y As Single) As Long
    Dim i                           As Long
    Dim j                           As Long
    Dim k                           As Long
    i = NoFraction(X / StdW1)
    j = NoFraction(Y / StdH1) * CharsInRow
    k = pageStart + j + i
    GetByteIndex = k
End Function
Private Function HexToBinStr(ByVal inHex As String) As String
    Dim mDec                        As Integer
    Dim s                           As String
    Dim i
    mDec = CInt("&h" & inHex)
    s = Trim(CStr(mDec Mod 2))
    i = mDec \ 2
    Do While i <> 0
        s = Trim(CStr(i Mod 2)) & s
        i = i \ 2
    Loop
    Do While Len(s) < 8
        s = "0" & s
    Loop
    HexToBinStr = s
    Exit Function
End Function
Public Sub picHexDisp_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If FrmMain.lblFileSpec.Caption = "" Then Exit Sub
    Dim k                           As Long
    Dim i                           As Long
    Dim j                           As Long
    Dim mHex                        As String
    k = GetByteIndex(X, Y)
    If k > pageEnd Then Exit Sub
    ShowPage True, k, k, vbBlue, vbYellow
    mHex = Hex$(arrByte(k))
    If Len(mHex) < 2 Then mHex = "0" & mHex
    FrmMain.lblAscii.Caption = Trim(CStr(CInt("&h" & mHex)))
    FrmMain.lblBinary.Caption = HexToBinStr(mHex)
End Sub
Public Sub picChrDisp_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If FrmMain.lblFileSpec.Caption = "" Then Exit Sub
    Dim i, j
    Dim k As Long
    Dim mHex As String
    i = NoFraction(X / StdW2)
    j = NoFraction(Y / StdH2) * CharsInRow
    k = pageStart + j + i
    If k > pageEnd Then Exit Sub
    ShowPage True, k, k, vbBlue, vbYellow
    mHex = Hex$(arrByte(k))
    If Len(mHex) < 2 Then mHex = "0" & mHex
    FrmMain.lblAscii.Caption = Trim(CStr(CInt("&h" & mHex)))
    FrmMain.lblBinary.Caption = HexToBinStr(mHex)
End Sub
Public Sub HexMovePageDown()
    If FrmMain.lblFileSpec.Caption = "" Then Exit Sub
    If mFileSize <= mPageSize Then Exit Sub
    If pageEnd = mFileSize Then Exit Sub
    FrmMain.picHexDisp.SetFocus
    pageStart = pageStart + mPageSize
    If pageStart > mFileSize Then pageStart = pageStart - mPageSize
    pageEnd = pageEnd + mPageSize
    If pageEnd > mFileSize Then pageEnd = mFileSize
    ShowPage False
End Sub
Public Sub HexMoveDown()
    If FrmMain.lblFileSpec.Caption = "" Then Exit Sub
    If mFileSize <= mPageSize Then Exit Sub
    FrmMain.picHexDisp.SetFocus
    pageStart = pageStart + CharsInRow
    If pageStart > mFileSize Then pageStart = pageStart - CharsInRow
    pageEnd = pageStart + mPageSize - 1
    If pageEnd > mFileSize Then pageEnd = mFileSize
    ShowPage False
End Sub
Public Sub HexMovePageUp()
    If FrmMain.lblFileSpec.Caption = "" Then Exit Sub
    If mFileSize <= mPageSize Then Exit Sub
    If pageStart = 1 Then Exit Sub
    FrmMain.picHexDisp.SetFocus
    pageStart = pageStart - mPageSize
    If pageStart < 1 Then pageStart = 1
    pageEnd = pageStart + mPageSize - 1
    If pageEnd > mFileSize Then pageEnd = mFileSize
    ShowPage False
End Sub
Public Sub HexMoveUp()
    If mFileSize <= mPageSize Then Exit Sub
    If pageStart <= CharsInRow Then Exit Sub
    FrmMain.picHexDisp.SetFocus
    pageStart = pageStart - CharsInRow
    If pageStart < 1 Then pageStart = pageStart + CharsInRow
    pageEnd = pageStart + mPageSize - 1
    If pageEnd > mFileSize Then pageEnd = mFileSize
    ShowPage False
End Sub

