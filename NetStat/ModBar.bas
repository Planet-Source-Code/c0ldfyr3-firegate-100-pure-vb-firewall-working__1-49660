Attribute VB_Name = "ModBar"
Option Explicit
'Niknak!! , http://www.Planet-Source-Code.com/vb/scripts/ShowCode.asp?txtCodeId=22974&lngWId=1
Private Declare Function CombineRgn Lib "gdi32" (ByVal hDestRgn As Long, ByVal hSrcRgn1 As Long, ByVal hSrcRgn2 As Long, ByVal nCombineMode As Long) As Long
Private Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function CreateRectRgn Lib "gdi32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function GetPixel Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Public Declare Function SetWindowRgn Lib "user32" (ByVal hwnd As Long, ByVal hRgn As Long, ByVal bRedraw As Boolean) As Long
Public Declare Sub ReleaseCapture Lib "user32" ()
Private Const RGN_OR                As Integer = 2
Private Const RGN_XOR               As Integer = 3
Private Const WM_NCLBUTTONDOWN      As Long = &HA1
Private Const HTCAPTION             As Integer = 2
Private FORM_WIDTH                  As Long    'FORM WIDTH
Private FORM_HEIGHT                 As Long    'FORM HEIGHT
Private DC                          As Long    'FORM DC HANDLE
Private BMP                         As Long    'BITMAP HANDLE
Private Pix                         As Long    'CURRENT PIXEL COLOUR
Private rgnInv                      As Long    'REGION JUNK
Private rgn                         As Long    'REGION JUNK
Private rgnTotal                    As Long    'REGION JUNK
    
Public Reshaping                    As Boolean
Public Reshape_Map                  As String
Public Function CreateRegionFromFile(Shape_Form As Form, shape_map As Image, strFile, BGColor As Long) As Long
    Dim height_counter              As Integer 'WIDTH COUNTER
    Dim width_counter               As Integer 'HEIGHT COUNTER
    shape_map.Picture = LoadPicture(strFile) 'LOAD THE PICTURE INTO AN IMAGE BOX ONTO A NON BORDERED FORM!
    Shape_Form.Width = shape_map.Width 'WRAP THE FORM TO THE IMAGE BOX
    Shape_Form.Height = shape_map.Height
    shape_map.Visible = False 'MAKE THE IMAGE INVISIBLE
    Shape_Form.Picture = LoadPicture(strFile) 'LOAD THE IMAGE INTO THE BACK OF THE FORM
    Shape_Form.ScaleMode = vbPixels 'SET THE SCALE MODE TO PIXELS
    'SET THE TWO VARIABLES
    'FW - FORM WIDTH
    'FH - FORM HEIGHT
    'TO THE WIDTH AND HEIGHT OF THE SCALED FORM
    FORM_WIDTH = Shape_Form.ScaleWidth
    FORM_HEIGHT = Shape_Form.ScaleHeight
    DC = CreateCompatibleDC(Shape_Form.hdc) 'CREATE COMPATIBLE DISPLAY CONTEXT
    BMP = SelectObject(DC, LoadPicture(strFile)) 'LOAD THE FORM BITMAP INTO BMP
    rgnTotal = CreateRectRgn(0, 0, FORM_WIDTH, FORM_HEIGHT) 'REGION SETUP
    rgnInv = CreateRectRgn(0, 0, FORM_WIDTH, FORM_HEIGHT)
    CombineRgn rgnTotal, rgnTotal, rgnTotal, RGN_XOR
    For height_counter = 0 To FORM_HEIGHT 'GO THROUGH THE FORM AND REMOVE ALL BACKGROUND COLOURED PIXELS
        For width_counter = 0 To FORM_WIDTH
            Pix = GetPixel(DC, width_counter, height_counter)
            If Pix = BGColor Then
                rgn = CreateRectRgn(width_counter, height_counter, width_counter + 1, height_counter + 1)
                CombineRgn rgnTotal, rgnTotal, rgn, RGN_OR
                DeleteObject rgn
            End If
        Next width_counter
    Next height_counter
    CombineRgn rgnTotal, rgnTotal, rgnInv, RGN_XOR
    CreateRegionFromFile = rgnTotal
End Function

