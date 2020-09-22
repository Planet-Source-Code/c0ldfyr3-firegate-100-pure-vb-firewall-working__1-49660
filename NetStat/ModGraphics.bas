Attribute VB_Name = "ModGraphics"
Option Explicit
Public Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Private Declare Function CreateBrushIndirect Lib "gdi32" (lpLogBrush As LogBrush) As Long
Private Declare Function CreateCompatibleBitmap Lib "gdi32" (ByVal hdc As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
Private Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function CreateFontIndirect Lib "gdi32" Alias "CreateFontIndirectA" (lpLogFont As LogFont) As Long
Private Declare Function CreatePenIndirect Lib "gdi32" (lpLogPen As LogPen) As Long
Private Declare Function DeleteDC Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function DrawText Lib "user32" Alias "DrawTextA" (ByVal hdc As Long, ByVal lpStr As String, ByVal nCount As Long, lpRect As RECT, ByVal wFormat As Long) As Long
Private Declare Function GetDeviceCaps Lib "gdi32" (ByVal hdc As Long, ByVal nIndex As Long) As Long
Private Declare Function GetLastError Lib "kernel32" () As Long
Private Declare Function GetTextMetrics Lib "gdi32" Alias "GetTextMetricsA" (ByVal hdc As Long, lpMetrics As TEXTMETRIC) As Long
Private Declare Function PatBlt Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal dwRop As Long) As Long
Private Declare Function Polyline Lib "gdi32" (ByVal hdc As Long, lpPoint As POINTAPI, ByVal nCount As Long) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
Private Declare Function SetBkColor Lib "gdi32" (ByVal hdc As Long, ByVal crColor As Long) As Long
Private Declare Function SetBkMode Lib "gdi32" (ByVal hdc As Long, ByVal nBkMode As Long) As Long
Private Declare Function SetTextColor Lib "gdi32" (ByVal hdc As Long, ByVal crColor As Long) As Long
Public Enum enBackStyle
    TRANSPARENT = 1
    OPAQUE = 2
End Enum
Public Enum AlignText
    vbLeftAlign = 1
    vbCentreAlign = 2
    vbRightAlign = 3
End Enum
Public Enum GradientTo
    GradHorizontal = 0
    GradVertical = 1
End Enum
Public Enum Scaling
    InTwips = 0
    InPixels = 1
End Enum
Private Enum en_PS_CONSTANTS 'pen constants
    PS_COSMETIC = &H0
    PS_DASH = 1                    '  -------
    PS_DASHDOT = 3                 '  _._._._
    PS_DASHDOTDOT = 4              '  _.._.._
    PS_DOT = 2                     '  .......
    PS_ENDCAP_ROUND = &H0
    PS_ENDCAP_SQUARE = &H100
    PS_ENDCAP_FLAT = &H200
    PS_GEOMETRIC = &H10000
    PS_INSIDEFRAME = 6
    PS_JOIN_BEVEL = &H1000
    PS_JOIN_MITER = &H2000
    PS_JOIN_ROUND = &H0
    PS_SOLID = 0
End Enum
Public Type FontStruc
    Name                            As String
    Alignment                       As AlignText
    Bold                            As Boolean
    Italic                          As Boolean
    Underline                       As Boolean
    StrikeThru                      As Boolean
    PointSize                       As Byte
    Colour                          As Long
End Type
Private Type LogPen
    lopnStyle                       As Long
    lopnWidth                       As POINTAPI
    lopnColor                       As Long
End Type
Private Type LogBrush
    lbStyle                         As Long
    lbColor                         As Long
    lbHatch                         As Long
End Type
Private Type RGBVal
    Red                             As Single
    Green                           As Single
    Blue                            As Single
End Type
Public Type BitmapStruc
    hDcMemory                       As Long
    hDcBitmap                       As Long
    hDcPointer                      As Long
    Area                            As RECT
End Type
Private Type LogFont
    'for the DrawText api call
    lfHeight                        As Long
    lfWidth                         As Long
    lfEscapement                    As Long
    lfOrientation                   As Long
    lfWeight                        As Long
    lfItalic                        As Byte
    lfUnderline                     As Byte
    lfStrikeOut                     As Byte
    lfCharSet                       As Byte
    lfOutPrecision                  As Byte
    lfClipPrecision                 As Byte
    lfQuality                       As Byte
    lfPitchAndFamily                As Byte
    lfFaceName(1 To 32)             As Byte
End Type
Private Type TEXTMETRIC
    tmHeight                        As Long
    tmAscent                        As Long
    tmDescent                       As Long
    tmInternalLeading               As Long
    tmExternalLeading               As Long
    tmAveCharWidth                  As Long
    tmMaxCharWidth                  As Long
    tmWeight                        As Long
    tmOverhang                      As Long
    tmDigitizedAspectX              As Long
    tmDigitizedAspectY              As Long
    tmFirstChar                     As Byte
    tmLastChar                      As Byte
    tmDefaultChar                   As Byte
    tmBreakChar                     As Byte
    tmItalic                        As Byte
    tmUnderlined                    As Byte
    tmStruckOut                     As Byte
    tmPitchAndFamily                As Byte
    tmCharSet                       As Byte
End Type
Public Const SRCAND                 As Long = &H8800C6   ' (DWORD) dest = source AND dest
Public Const SRCCOPY                As Long = &HCC0020   ' (DWORD) dest = source
Public Const SRCERASE               As Long = &H440328   ' (DWORD) dest = source AND (NOT dest )
Public Const SRCINVERT              As Long = &H660046   ' (DWORD) dest = source XOR dest
Public Const SRCPAINT               As Long = &HEE0086   ' (DWORD) dest = source OR dest
Public Const MERGECOPY              As Long = &HC000CA   ' (DWORD) dest = (source AND pattern)
Public Const MERGEPAINT             As Long = &HBB0226   ' (DWORD) dest = (NOT source) OR dest
Public Const NOTSRCCOPY             As Long = &H330008   ' (DWORD) dest = (NOT source)
Public Const NOTSRCERASE            As Long = &H1100A6   ' (DWORD) dest = (NOT src) AND (NOT dest)
Private Const BS_DIBPATTERN         As Long = 5
Private Const BS_DIBPATTERN8X8      As Long = 8
Private Const BS_DIBPATTERNPT       As Long = 6
Private Const BS_HATCHED            As Long = 2
Private Const BS_HOLLOW             As Long = 1
Private Const BS_NULL               As Long = 1
Private Const BS_PATTERN            As Long = 3
Private Const BS_PATTERN8X8         As Long = 7
Private Const BS_SOLID              As Long = 0
Private Const HS_BDIAGONAL          As Long = 3          '  /////
Private Const HS_CROSS              As Long = 4          '  +++++
Private Const HS_DIAGCROSS          As Long = 5          '  xxxxx
Private Const HS_FDIAGONAL          As Long = 2          '  \\\\\
Private Const HS_HORIZONTAL         As Long = 0          '  -----
Private Const HS_NOSHADE            As Long = 17
Private Const HS_SOLID              As Long = 8
Private Const HS_SOLIDBKCLR         As Long = 23
Private Const HS_SOLIDCLR           As Long = 19
Private Const HS_VERTICAL           As Long = 1
Private Const LOGPIXELSY            As Long = 90         '  Logical pixels/inch in Y
Private Const LF_FACESIZE           As Long = 32
Private Const FW_BOLD               As Long = 700
Private Const FW_DONTCARE           As Long = 0
Private Const FW_EXTRABOLD          As Long = 800
Private Const FW_EXTRALIGHT         As Long = 200
Private Const FW_HEAVY              As Long = 900
Private Const FW_LIGHT              As Long = 300
Private Const FW_MEDIUM             As Long = 500
Private Const FW_NORMAL             As Long = 400
Private Const FW_SEMIBOLD           As Long = 600
Private Const FW_THIN               As Long = 100
Private Const DEFAULT_CHARSET       As Long = 1
Private Const OUT_CHARACTER_PRECIS  As Long = 2
Private Const OUT_DEFAULT_PRECIS    As Long = 0
Private Const OUT_DEVICE_PRECIS     As Long = 5
Private Const OUT_OUTLINE_PRECIS    As Long = 8
Private Const OUT_RASTER_PRECIS     As Long = 6
Private Const OUT_STRING_PRECIS     As Long = 1
Private Const OUT_STROKE_PRECIS     As Long = 3
Private Const OUT_TT_ONLY_PRECIS    As Long = 7
Private Const OUT_TT_PRECIS         As Long = 4
Private Const CLIP_CHARACTER_PRECIS As Long = 1
Private Const CLIP_DEFAULT_PRECIS   As Long = 0
Private Const CLIP_EMBEDDED         As Long = 128
Private Const CLIP_LH_ANGLES        As Long = 16
Private Const CLIP_MASK             As Long = &HF
Private Const CLIP_STROKE_PRECIS    As Long = 2
Private Const CLIP_TT_ALWAYS        As Long = 32
Private Const WM_SETFONT            As Long = &H30
Private Const LF_FULLFACESIZE       As Long = 64
Private Const DEFAULT_PITCH         As Long = 0
Private Const DEFAULT_QUALITY       As Long = 0
Private Const PROOF_QUALITY         As Long = 2
Private Const DT_CENTER             As Long = &H1
Private Const DT_BOTTOM             As Long = &H8
Private Const DT_CALCRECT           As Long = &H400
Private Const DT_EXPANDTABS         As Long = &H40
Private Const DT_EXTERNALLEADING    As Long = &H200
Private Const DT_LEFT               As Long = &H0
Private Const DT_NOCLIP             As Long = &H100
Private Const DT_NOPREFIX           As Long = &H800
Private Const DT_RIGHT              As Long = &H2
Private Const DT_SINGLELINE         As Long = &H20
Private Const DT_TABSTOP            As Long = &H80
Private Const DT_TOP                As Long = &H0
Private Const DT_VCENTER            As Long = &H4
Private Const DT_WORDBREAK          As Long = &H10
Private Const PATCOPY               As Long = &HF00021   ' (DWORD) dest = pattern
Private Const PATINVERT             As Long = &H5A0049   ' (DWORD) dest = pattern XOR dest
Private Const PATPAINT              As Long = &HFB0A09   ' (DWORD) dest = DPSnoo
Private Const DSTINVERT             As Long = &H550009   ' (DWORD) dest = (NOT dest)
Private Const BLACKNESS             As Long = &H42       ' (DWORD) dest = BLACK
Private Const WHITENESS             As Long = &HFF0062   ' (DWORD) dest = WHITE
Public Function GetTextHeight(ByVal hdc As Long) As Integer
    'This function will return the height of the text using the point size
    Dim udtMetrics                  As TEXTMETRIC
    Dim lngResult                   As Long
    lngResult = GetTextMetrics(hdc, udtMetrics)
    GetTextHeight = udtMetrics.tmHeight
End Function
Public Sub CreateNewBitmap(ByRef hDcMemory As Long, ByRef hDcBitmap As Long, ByRef hDcPointer As Long, ByRef BmpArea As RECT, ByVal CompatableWithhDc As Long, Optional ByVal lngBackColour As Long = 0, Optional ByVal udtMeasurement As Scaling = InPixels)
    'This procedure will create a new bitmap compatable with a given
    'form (you will also be able to then use this bitmap in a picturebox).
    'The space specified in "Area" should be in "Twips" and will be
    'converted into pixels in the following code.
    Dim lngResult                   As Long
    Dim Area                        As RECT
    'scale the bitmap points if necessary
    Area = BmpArea
    If udtMeasurement = InTwips Then Call RectToPixels(Area)
    'create the bitmap and its references
    hDcMemory = CreateCompatibleDC(CompatableWithhDc)
    hDcBitmap = CreateCompatibleBitmap(CompatableWithhDc, (Area.Right - Area.Left), (Area.Bottom - Area.Top))
    hDcPointer = SelectObject(hDcMemory, hDcBitmap)
    'set default colours and clear bitmap to selected colour
    lngResult = SetBkMode(hDcMemory, OPAQUE)
    lngResult = SetBkColor(hDcMemory, lngBackColour)
    Call DrawRect(hDcMemory, lngBackColour, 0, 0, (Area.Right - Area.Left), (Area.Bottom - Area.Top))
End Sub
Public Sub Gradient(ByVal lngDesthDc As Long, ByVal lngStartCol As Long, ByVal FinishCol As Long, ByVal intLeft As Integer, ByVal intTop As Integer, ByVal intWidth As Integer, ByVal intHeight As Integer, ByVal Direction As GradientTo, Optional ByVal udtMeasurement As Scaling = 1, Optional ByVal bytLineWidth As Byte = 1)
    'draws a gradient from colour mblnStart to colour Finish, and assums
    'that all measurments passed to it are in pixels unless otherwise
    'specified.
    Dim intCounter                  As Integer
    Dim intBiggestDiff              As Integer
    Dim Colour                      As RGBVal
    Dim mblnStart                   As RGBVal
    Dim Finish                      As RGBVal
    Dim sngAddRed                   As Single
    Dim sngAddGreen                 As Single
    Dim sngAddBlue                  As Single
    'perform all necessary calculations before drawing gradient
    'such as converting long to rgb values, and getting the correct
    'scaling for the bitmap.
    mblnStart = GetRGB(lngStartCol)
    Finish = GetRGB(FinishCol)
    If udtMeasurement = InTwips Then
        intLeft = intLeft / Screen.TwipsPerPixelX
        intTop = intTop / Screen.TwipsPerPixelY
        intWidth = intWidth / Screen.TwipsPerPixelX
        intHeight = intHeight / Screen.TwipsPerPixelY
    End If
    'draw the colour gradient
    Select Case Direction
        Case GradVertical
            intBiggestDiff = intWidth
        Case GradHorizontal
            intBiggestDiff = intHeight
    End Select
    'calculate how much to increment/decrement each colour per step
    sngAddRed = (bytLineWidth * ((Finish.Red) - mblnStart.Red) / intBiggestDiff)
    sngAddGreen = (bytLineWidth * ((Finish.Green) - mblnStart.Green) / intBiggestDiff)
    sngAddBlue = (bytLineWidth * ((Finish.Blue) - mblnStart.Blue) / intBiggestDiff)
    Colour = mblnStart
    'calculate the colour of each line before drawing it on the bitmap
    For intCounter = 0 To intBiggestDiff Step bytLineWidth
        'find the point between colour mblnStart and Colour Finish that
        'corresponds to the point between 0 and intBiggestDiff
        'check for overflow
        If Colour.Red > 255 Then
            Colour.Red = 255
        Else
            If Colour.Red < 0 Then Colour.Red = 0
        End If
        If Colour.Green > 255 Then
            Colour.Green = 255
        Else
            If Colour.Green < 0 Then Colour.Green = 0
        End If
        If Colour.Blue > 255 Then
            Colour.Blue = 255
        Else
            If Colour.Blue < 0 Then Colour.Blue = 0
        End If
        'draw the gradient in the proper orientation in the calculated colour
        Select Case Direction
            Case GradVertical
                Call DrawLine(lngDesthDc, intCounter + intLeft, intTop, intCounter + intLeft, intHeight + intTop, RGB(Colour.Red, Colour.Green, Colour.Blue), bytLineWidth, InPixels)
            Case GradHorizontal
                Call DrawLine(lngDesthDc, intLeft, intCounter + intTop, intLeft + intWidth, intTop + intCounter, RGB(Colour.Red, Colour.Green, Colour.Blue), bytLineWidth, InPixels)
        End Select
        'set next colour
        Colour.Red = Colour.Red + sngAddRed
        Colour.Green = Colour.Green + sngAddGreen
        Colour.Blue = Colour.Blue + sngAddBlue
    Next intCounter
End Sub
Public Sub MakeText(ByVal hDcSurphase As Long, ByVal strText As String, ByVal intTop As Integer, ByVal intLeft As Integer, ByVal intHeight As Integer, ByVal intWidth As Integer, ByRef udtFont As FontStruc, Optional ByVal udtMeasurement As Scaling = 0)
    'This procedure will draw strText onto the bitmap in the specified udtFont,
    'colour and position.
    Dim udtAPIFont                  As LogFont
    Dim lngAlignment                As Long
    Dim udtTextRect                 As RECT
    Dim lngResult                   As Long
    Dim lngJunk                     As Long
    Dim hDcFont                     As Long
    Dim hDcOldFont                  As Long
    Dim intCounter                  As Integer
    'set Measurement values
    udtTextRect.Top = intTop
    udtTextRect.Left = intLeft
    udtTextRect.Right = intLeft + intWidth
    udtTextRect.Bottom = intTop + intHeight
    If udtMeasurement = InTwips Then Call RectToPixels(udtTextRect) 'convert to pixels
    'Create details about the udtFont using the udtFont structure
    '====================
    'convert point size to pixels
    udtAPIFont.lfHeight = -((udtFont.PointSize * GetDeviceCaps(hDcSurphase, LOGPIXELSY)) / 72)
    udtAPIFont.lfCharSet = DEFAULT_CHARSET
    udtAPIFont.lfClipPrecision = CLIP_DEFAULT_PRECIS
    udtAPIFont.lfEscapement = 0
    'move the name of the udtFont into the array
    For intCounter = 1 To Len(udtFont.Name)
        udtAPIFont.lfFaceName(intCounter) = Asc(Mid(udtFont.Name, intCounter, 1))
    Next intCounter
    'this has to be a Null terminated string
    udtAPIFont.lfFaceName(intCounter) = 0
    udtAPIFont.lfItalic = udtFont.Italic
    udtAPIFont.lfUnderline = udtFont.Underline
    udtAPIFont.lfStrikeOut = udtFont.StrikeThru
    udtAPIFont.lfOrientation = 0
    udtAPIFont.lfOutPrecision = OUT_DEFAULT_PRECIS
    udtAPIFont.lfPitchAndFamily = DEFAULT_PITCH
    udtAPIFont.lfQuality = PROOF_QUALITY
    udtAPIFont.lfWeight = IIf(udtFont.Bold, FW_BOLD, FW_NORMAL)
    udtAPIFont.lfWidth = 0
    hDcFont = CreateFontIndirect(udtAPIFont)
    hDcOldFont = SelectObject(hDcSurphase, hDcFont)
    '====================
    Select Case udtFont.Alignment
        Case vbLeftAlign
            lngAlignment = DT_LEFT
        Case vbCentreAlign
            lngAlignment = DT_CENTER
        Case vbRightAlign
            lngAlignment = DT_RIGHT
    End Select
    'Draw the strText into the off-screen bitmap before copying the
    'new bitmap (with the strText) onto the screen.
    lngResult = SetBkMode(hDcSurphase, TRANSPARENT)
    lngResult = SetTextColor(hDcSurphase, udtFont.Colour)
    lngResult = DrawText(hDcSurphase, strText, Len(strText), udtTextRect, lngAlignment)
    'clean up by deleting the off-screen bitmap and udtFont
    lngJunk = SelectObject(hDcSurphase, hDcOldFont)
    lngJunk = DeleteObject(hDcFont)
End Sub
Public Sub DeleteBitmap(ByRef hDcMemory As Long, ByRef hDcBitmap As Long, ByRef hDcPointer As Long)
    'This will remove the bitmap that stored what was displayed before
    'the text was written to the screen, from memory.
    Dim lngJunk                     As Long
    If hDcMemory = 0 Then Exit Sub 'there is nothing to delete. Exit the sub-routine
    'delete the device context
    lngJunk = SelectObject(hDcMemory, hDcPointer)
    lngJunk = DeleteObject(hDcBitmap)
    lngJunk = DeleteDC(hDcMemory)
    'show that the device context has been deleted by setting
    'all parameters passed to the procedure to zero
    hDcMemory = 0
    hDcBitmap = 0
    hDcPointer = 0
End Sub
Public Sub Pause(lngTicks As Long)
    Dim J_Timer                     As clsWaitableTimer
    'pause execution of the program for a specified number of lngTicks
    Set J_Timer = New clsWaitableTimer
    If lngTicks < 0 Then lngTicks = 0
    J_Timer.Wait lngTicks
    Set J_Timer = Nothing
End Sub
Public Sub RectToPixels(ByRef TheRect As RECT)
    'converts twips to pixels in a rect structure
    TheRect.Left = TheRect.Left \ Screen.TwipsPerPixelX
    TheRect.Right = TheRect.Right \ Screen.TwipsPerPixelX
    TheRect.Top = TheRect.Top \ Screen.TwipsPerPixelY
    TheRect.Bottom = TheRect.Bottom \ Screen.TwipsPerPixelY
End Sub
Public Sub DrawRect(ByVal lngHDC As Long, ByVal lngColour As Long, ByVal intLeft As Integer, ByVal intTop As Integer, ByVal intRight As Integer, ByVal intBottom As Integer, Optional ByVal udtMeasurement As Scaling = InPixels, Optional ByVal lngStyle As Long = BS_SOLID, Optional ByVal lngPattern As Long = HS_SOLID)
    'this draws a rectangle using the co-ordinates
    'and lngColour given.
    Dim StartRect                   As RECT
    Dim lngResult                   As Long
    Dim lngJunk                     As Long
    Dim lnghBrush                   As Long
    Dim BrushStuff                  As LogBrush
    'check if conversion is necessary
    If udtMeasurement = InTwips Then
        'convert to pixels
        intLeft = intLeft / Screen.TwipsPerPixelX
        intTop = intTop / Screen.TwipsPerPixelY
        intRight = intRight / Screen.TwipsPerPixelX
        intBottom = intBottom / Screen.TwipsPerPixelY
    End If
    'initalise values
    StartRect.Top = intTop
    StartRect.Left = intLeft
    StartRect.Bottom = intBottom
    StartRect.Right = intRight
    'create a brush
    BrushStuff.lbColor = lngColour
    BrushStuff.lbHatch = lngPattern
    BrushStuff.lbStyle = lngStyle
    'apply the brush to the device context
    lnghBrush = CreateBrushIndirect(BrushStuff)
    lnghBrush = SelectObject(lngHDC, lnghBrush)
    'draw a rectangle
    lngResult = PatBlt(lngHDC, intLeft, intTop, (intRight - intLeft), (intBottom - intTop), PATCOPY)
    'A "Brush" object was created. It must be removed from memory.
    lngJunk = SelectObject(lngHDC, lnghBrush)
    lngJunk = DeleteObject(lngJunk)
End Sub
Public Function GetRGB(ByVal lngColour As Long) As RGBVal
    'Convert Long to RGB:
    'if the lngcolour value is greater than acceptable then half the value
    If (lngColour > RGB(255, 255, 255)) Or (lngColour < (RGB(255, 255, 255) * -1)) Then Exit Function
    GetRGB.Blue = (lngColour \ 65536)
    GetRGB.Green = ((lngColour - (GetRGB.Blue * 65536)) \ 256)
    GetRGB.Red = (lngColour - (GetRGB.Blue * (65536)) - ((GetRGB.Green) * 256))
End Function
Public Sub DrawLine(lngHDC As Long, ByVal intX1 As Integer, ByVal intY1 As Integer, ByVal intX2 As Integer, ByVal intY2 As Integer, Optional ByVal lngColour As Long = 0, Optional ByVal intWidth As Integer = 1, Optional ByVal udtMeasurement As Scaling = InTwips)
    'This will draw a line from point1 to point2
    Const NumOfPoints = 2
    Dim lngResult                   As Long
    Dim lnghPen                     As Long
    Dim PenStuff                    As LogPen
    Dim Junk                        As Long
    Dim Points(NumOfPoints)         As POINTAPI
    'check if conversion is necessary
    If udtMeasurement = InTwips Then
        'convert twip values to pixels
        intX1 = intX1 / Screen.TwipsPerPixelX
        intX2 = intX2 / Screen.TwipsPerPixelX
        intY1 = intY1 / Screen.TwipsPerPixelY
        intY2 = intY2 / Screen.TwipsPerPixelY
    End If
    'Find out if a specific lngColour is to be set. If so set it.
    PenStuff.lopnColor = lngColour
    PenStuff.lopnStyle = PS_GEOMETRIC
    PenStuff.lopnWidth.X = intWidth
    'apply the pen settings to the device context
    lnghPen = CreatePenIndirect(PenStuff)
    lnghPen = SelectObject(lngHDC, lnghPen)
    'set the points
    Points(1).X = intX1
    Points(1).Y = intY1
    Points(2).X = intX2
    Points(2).Y = intY2
    'draw the line
    lngResult = Polyline(lngHDC, Points(1), NumOfPoints)
    lngResult = GetLastError
    'A "Pen" object was created. It must be removed from memory.
    Junk = SelectObject(lngHDC, lnghPen)
    Junk = DeleteObject(Junk)
End Sub
