Attribute VB_Name = "ModCoolMenu"
Option Explicit

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (hpvDest As Any, hpvSource As Any, ByVal cbCopy As Long)
Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Private Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hwnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function ChildWindowFromPoint Lib "user32" (ByVal hWndParent As Long, pt As POINTAPI) As Long
Private Declare Function CreateBitmap Lib "gdi32" (ByVal nWidth As Long, ByVal nHeight As Long, ByVal nPlanes As Long, ByVal nBitCount As Long, lpBits As Any) As Long
Private Declare Function CreateBrushIndirect Lib "gdi32" (lpLogBrush As LogBrush) As Long
Private Declare Function CreateCompatibleBitmap Lib "gdi32" (ByVal hdc As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
Private Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function CreateFontIndirect Lib "gdi32" Alias "CreateFontIndirectA" (lpLogFont As LogFont) As Long
Private Declare Function CreateHatchBrush Lib "gdi32" (ByVal nIndex As Long, ByVal crColor As Long) As Long
Private Declare Function CreatePatternBrush Lib "gdi32" (ByVal hBitmap As Long) As Long
Private Declare Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Long) As Long
Private Declare Function DeleteDC Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function DrawEdge Lib "user32" (ByVal hdc As Long, qrc As RECT, ByVal edge As Long, ByVal grfFlags As Long) As Long
Private Declare Function DrawMenuBar Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function DrawState Lib "user32" Alias "DrawStateA" (ByVal hdc As Long, ByVal hBrush As Long, ByVal lpDrawStateProc As Long, ByVal lParam As Long, ByVal wParam As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal fuFlags As Long) As Long
Private Declare Function DrawText Lib "user32" Alias "DrawTextA" (ByVal hdc As Long, ByVal lpStr As String, ByVal nCount As Long, lpRect As RECT, ByVal wFormat As Long) As Long
Private Declare Function ExcludeClipRect Lib "gdi32" (ByVal hdc As Long, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Declare Function FillRect Lib "user32" (ByVal hdc As Long, lpRect As RECT, ByVal hBrush As Long) As Long
Private Declare Function FormatMessage Lib "kernel32" Alias "FormatMessageA" (ByVal dwFlags As Long, lpSource As Any, ByVal dwMessageId As Long, ByVal dwLanguageId As Long, ByVal lpBuffer As String, ByVal nSize As Long, Arguments As Long) As Long
Private Declare Function GetClientRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long
Private Declare Function GetDC Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function GetDeviceCaps Lib "gdi32" (ByVal hdc As Long, ByVal nIndex As Long) As Long
Private Declare Function GetLastError Lib "kernel32" () As Long
Private Declare Function GetMenu Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function GetMenuItemCount Lib "user32" (ByVal hMenu As Long) As Long
Private Declare Function GetMenuItemID Lib "user32" (ByVal hMenu As Long, ByVal nPos As Long) As Long
Private Declare Function GetMenuItemInfo Lib "user32" Alias "GetMenuItemInfoA" (ByVal hMenu As Long, ByVal uItem As Long, ByVal ByPosition As Long, ByRef lpMenuItemInfo As MENUITEMINFO) As Boolean
Private Declare Function GetMenuItemRect Lib "user32" (ByVal hwnd As Long, ByVal hMenu As Long, ByVal uItem As Long, lprcItem As RECT) As Long
Private Declare Function GetMenuState Lib "user32" (ByVal hMenu As Long, ByVal wID As Long, ByVal wFlags As Long) As Long
Private Declare Function GetParent Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function GetSubMenu Lib "user32" (ByVal hMenu As Long, ByVal nPos As Long) As Long
Private Declare Function GetSysColor Lib "user32" (ByVal nIndex As Long) As Long
Private Declare Function GetSystemMetrics Lib "user32" (ByVal nIndex As Long) As Long
Private Declare Function GetTextColor Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function GetTextFace Lib "gdi32" Alias "GetTextFaceA" (ByVal hdc As Long, ByVal nCount As Long, ByVal lpFacename As String) As Long
Private Declare Function GetTextMetrics Lib "gdi32" Alias "GetTextMetricsA" (ByVal hdc As Long, lpMetrics As TEXTMETRIC) As Long
Private Declare Function GetVersionEx Lib "kernel32" Alias "GetVersionExA" (lpVersionInformation As OSVERSIONINFO) As Long
Private Declare Function GetWindowDC Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Private Declare Function GetWindowRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long
Private Declare Function LoadImage Lib "user32" Alias "LoadImageA" (ByVal hinst As Long, ByVal lpszName As Any, ByVal uType As Long, ByVal cxDesired As Long, ByVal cyDesired As Long, ByVal fuLoad As Long) As Long
Private Declare Function MenuItemFromPoint Lib "user32" (ByVal hwnd As Long, ByVal hMenu As Long, ByVal ptScreen As Double) As Long
Private Declare Function MulDiv Lib "kernel32" (ByVal nNumber As Long, ByVal nNumerator As Long, ByVal nDenominator As Long) As Long
Private Declare Function PatBlt Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal dwRop As Long) As Long
Private Declare Function PostMessage Lib "user32" Alias "PostMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function ReleaseDC Lib "user32" (ByVal hwnd As Long, ByVal hdc As Long) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function SetBkColor Lib "gdi32" (ByVal hdc As Long, ByVal crColor As Long) As Long
Private Declare Function SetBkMode Lib "gdi32" (ByVal hdc As Long, ByVal nBkMode As Long) As Long
Private Declare Function SetMenuItemInfo Lib "user32" Alias "SetMenuItemInfoA" (ByVal hMenu As Long, ByVal un As Long, ByVal bool As Boolean, lpcMenuItemInfo As MENUITEMINFO) As Long
Private Declare Function SetTextColor Lib "gdi32" (ByVal hdc As Long, ByVal crColor As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function WindowFromDC Lib "user32" (ByVal hdc As Long) As Long
Public Declare Function DrawIconEx Lib "user32" (ByVal hdc As Long, ByVal xLeft As Long, ByVal yTop As Long, ByVal hIcon As Long, ByVal cxWidth As Long, ByVal cyWidth As Long, ByVal istepIfAniCur As Long, ByVal hbrFlickerFreeDraw As Long, ByVal diFlags As Long) As Long
Private Declare Function DrawIcon Lib "user32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal hIcon As Long) As Long
Private Declare Function CopyImage Lib "user32" (ByVal Handle As Long, ByVal un1 As Long, ByVal n1 As Long, ByVal n2 As Long, ByVal un2 As Long) As Long
Public Declare Function DestroyIcon Lib "user32" (ByVal hIcon As Long) As Long
Private Declare Function ImageList_GetIcon Lib "comctl32.dll" (ByVal himl As Long, ByVal i As Long, ByVal diIgnore As Long) As Long
Private Declare Function ImageList_Draw Lib "comctl32.dll" (ByVal himl As Long, ByVal i As Long, ByVal hdcDst As Long, ByVal X As Long, ByVal Y As Long, ByVal fStyle As Long) As Long
Private Declare Function ImageList_GetImageInfo Lib "comctl32.dll" (ByVal himl As Long, ByVal i As Long, IMAGEINFO As Any) As Long


Private Type LogBrush
  lbStyle                       As Long
  lbColor                       As Long
  lbHatch                       As Long
End Type


Private Const BS_SOLID = 0
Private Const BS_NULL = 1
Private Const BS_HOLLOW = BS_NULL
Private Const BS_HATCHED = 2
Private Const BS_PATTERN = 3
Private Const BS_INDEXED = 4
Private Const BS_DIBPATTERN = 5
Private Const BS_DIBPATTERNPT = 6
Private Const BS_PATTERN8X8 = 7
Private Const BS_DIBPATTERN8X8 = 8


Private Const IMAGE_BITMAP = 0&
Private Const IMAGE_ICON = 1&
Private Const IMAGE_CURSOR = 2&


Private Const LR_DEFAULTCOLOR = &H0
Private Const LR_MONOCHROME = &H1
Private Const LR_COLOR = &H2
Private Const LR_COPYRETURNORG = &H4
Private Const LR_COPYDELETEORG = &H8
Private Const LR_LOADFROMFILE = &H10
Private Const LR_LOADTRANSPARENT = &H20
Private Const LR_DEFAULTSIZE = &H40
Private Const LR_VGACOLOR = &H80
Private Const LR_LOADMAP3DCOLORS = &H1000
Private Const LR_CREATEDIBSECTION = &H2000
Private Const LR_COPYFROMRESOURCE = &H4000
Private Const LR_SHARED = &H8000


Private Const OBM_LFARROWI = 32734
Private Const OBM_RGARROWI = 32735
Private Const OBM_DNARROWI = 32736
Private Const OBM_UPARROWI = 32737
Private Const OBM_COMBO = 32738
Private Const OBM_MNARROW = 32739
Private Const OBM_LFARROWD = 32740
Private Const OBM_RGARROWD = 32741
Private Const OBM_DNARROWD = 32742
Private Const OBM_UPARROWD = 32743
Private Const OBM_RESTORED = 32744
Private Const OBM_ZOOMD = 32745
Private Const OBM_REDUCED = 32746
Private Const OBM_RESTORE = 32747
Private Const OBM_ZOOM = 32748
Private Const OBM_REDUCE = 32749
Private Const OBM_LFARROW = 32750
Private Const OBM_RGARROW = 32751
Private Const OBM_DNARROW = 32752
Private Const OBM_UPARROW = 32753
Private Const OBM_CLOSE = 32754
Private Const OBM_OLD_RESTORE = 32755
Private Const OBM_OLD_ZOOM = 32756
Private Const OBM_OLD_REDUCE = 32757
Private Const OBM_BTNCORNERS = 32758
Private Const OBM_CHECKBOXES = 32759
Private Const OBM_CHECK = 32760
Private Const OBM_BTSIZE = 32761
Private Const OBM_OLD_LFARROW = 32762
Private Const OBM_OLD_RGARROW = 32763
Private Const OBM_OLD_DNARROW = 32764
Private Const OBM_OLD_UPARROW = 32765
Private Const OBM_SIZE = 32766
Private Const OBM_OLD_CLOSE = 32767


Private Const SM_CXSCREEN = 0
Private Const SM_CYSCREEN = 1
Private Const SM_CXVSCROLL = 2
Private Const SM_CYHSCROLL = 3
Private Const SM_CYCAPTION = 4
Private Const SM_CXBORDER = 5
Private Const SM_CYBORDER = 6
Private Const SM_CXDLGFRAME = 7
Private Const SM_CYDLGFRAME = 8
Private Const SM_CYVTHUMB = 9
Private Const SM_CXHTHUMB = 10
Private Const SM_CXICON = 11
Private Const SM_CYICON = 12
Private Const SM_CXCURSOR = 13
Private Const SM_CYCURSOR = 14
Private Const SM_CYMENU = 15
Private Const SM_CXFULLSCREEN = 16
Private Const SM_CYFULLSCREEN = 17
Private Const SM_CYKANJIWINDOW = 18
Private Const SM_MOUSEPRESENT = 19
Private Const SM_CYVSCROLL = 20
Private Const SM_CXHSCROLL = 21
Private Const SM_DEBUG = 22
Private Const SM_SWAPBUTTON = 23
Private Const SM_RESERVED1 = 24
Private Const SM_RESERVED2 = 25
Private Const SM_RESERVED3 = 26
Private Const SM_RESERVED4 = 27
Private Const SM_CXMIN = 28
Private Const SM_CYMIN = 29
Private Const SM_CXSIZE = 30
Private Const SM_CYSIZE = 31
Private Const SM_CXFRAME = 32
Private Const SM_CYFRAME = 33
Private Const SM_CXMINTRACK = 34
Private Const SM_CYMINTRACK = 35
Private Const SM_CXDOUBLECLK = 36
Private Const SM_CYDOUBLECLK = 37
Private Const SM_CXICONSPACING = 38
Private Const SM_CYICONSPACING = 39
Private Const SM_MENUDROPALIGNMENT = 40
Private Const SM_PENWINDOWS = 41
Private Const SM_DBCSENABLED = 42
Private Const SM_CMOUSEBUTTONS = 43

Private Const SM_CXFIXEDFRAME = SM_CXDLGFRAME
Private Const SM_CYFIXEDFRAME = SM_CYDLGFRAME
Private Const SM_CXSIZEFRAME = SM_CXFRAME
Private Const SM_CYSIZEFRAME = SM_CYFRAME

Private Const SM_SECURE = 44
Private Const SM_CXEDGE = 45
Private Const SM_CYEDGE = 46
Private Const SM_CXMINSPACING = 47
Private Const SM_CYMINSPACING = 48
Private Const SM_CXSMICON = 49
Private Const SM_CYSMICON = 50
Private Const SM_CYSMCAPTION = 51
Private Const SM_CXSMSIZE = 52
Private Const SM_CYSMSIZE = 53
Private Const SM_CXMENUSIZE = 54
Private Const SM_CYMENUSIZE = 55
Private Const SM_ARRANGE = 56
Private Const SM_CXMINIMIZED = 57
Private Const SM_CYMINIMIZED = 58
Private Const SM_CXMAXTRACK = 59
Private Const SM_CYMAXTRACK = 60
Private Const SM_CXMAXIMIZED = 61
Private Const SM_CYMAXIMIZED = 62
Private Const SM_NETWORK = 63
Private Const SM_CLEANBOOT = 67
Private Const SM_CXDRAG = 68
Private Const SM_CYDRAG = 69
Private Const SM_SHOWSOUNDS = 70
Private Const SM_CXMENUCHECK = 71
Private Const SM_CYMENUCHECK = 72
Private Const SM_SLOWMACHINE = 73
Private Const SM_MIDEASTENABLED = 74


Private Const NULLREGION = 1
Private Const SIMPLEREGION = 2
Private Const COMPLEXREGION = 3


Private Const HS_HORIZONTAL = 0
Private Const HS_VERTICAL = 1
Private Const HS_FDIAGONAL = 2
Private Const HS_BDIAGONAL = 3
Private Const HS_CROSS = 4
Private Const HS_DIAGCROSS = 5
Private Const HS_FDIAGONAL1 = 6
Private Const HS_BDIAGONAL1 = 7
Private Const HS_SOLID = 8
Private Const HS_DENSE1 = 9
Private Const HS_DENSE2 = 10
Private Const HS_DENSE3 = 11
Private Const HS_DENSE4 = 12
Private Const HS_DENSE5 = 13
Private Const HS_DENSE6 = 14
Private Const HS_DENSE7 = 15
Private Const HS_DENSE8 = 16
Private Const HS_NOSHADE = 17
Private Const HS_HALFTONE = 18
Private Const HS_SOLIDCLR = 19
Private Const HS_DITHEREDCLR = 20
Private Const HS_SOLIDTEXTCLR = 21
Private Const HS_DITHEREDTEXTCLR = 22
Private Const HS_SOLIDBKCLR = 23
Private Const HS_DITHEREDBKCLR = 24
Private Const HS_API_MAX = 25


Private Const ILD_NORMAL = &H0
Private Const ILD_TRANSPARENT = &H1
Private Const ILD_MASK = &H10
Private Const ILD_IMAGE = &H20


Private Const DST_COMPLEX = &H0
Private Const DST_TEXT = &H1
Private Const DST_PREFIXTEXT = &H2
Private Const DST_ICON = &H3
Private Const DST_BITMAP = &H4


Private Const DSS_NORMAL = &H0
Private Const DSS_UNION = &H10
Private Const DSS_DISABLED = &H20
Private Const DSS_MONO = &H80
Private Const DSS_RIGHT = &H8000


Private Const COLOR_ACTIVEBORDER = 10
Private Const COLOR_ACTIVECAPTION = 2
Private Const COLOR_ADJ_MAX = 100
Private Const COLOR_ADJ_MIN = -100
Private Const COLOR_APPWORKSPACE = 12
Private Const COLOR_BACKGROUND = 1
Private Const COLOR_BTNFACE = 15
Private Const COLOR_BTNHIGHLIGHT = 20
Private Const COLOR_BTNSHADOW = 16
Private Const COLOR_BTNTEXT = 18
Private Const COLOR_CAPTIONTEXT = 9
Private Const COLOR_GRAYTEXT = 17
Private Const COLOR_HIGHLIGHT = 13
Private Const COLOR_HIGHLIGHTTEXT = 14
Private Const COLOR_INACTIVEBORDER = 11
Private Const COLOR_INACTIVECAPTION = 3
Private Const COLOR_INACTIVECAPTIONTEXT = 19
Private Const COLOR_BTNDKSHADOW = 21
Private Const COLOR_BTNLIGHT = 22
Private Const COLOR_MENU = 4
Private Const COLOR_MENUTEXT = 7
Private Const COLOR_SCROLLBAR = 0
Private Const COLOR_WINDOW = 5
Private Const COLOR_WINDOWFRAME = 6
Private Const COLOR_WINDOWTEXT = 8


Private Const ODA_DRAWENTIRE = &H1
Private Const ODA_SELECT = &H2
Private Const ODA_FOCUS = &H4


Private Const ODS_SELECTED = &H1
Private Const ODS_GRAYED = &H2
Private Const ODS_DISABLED = &H4
Private Const ODS_CHECKED = &H8
Private Const ODS_FOCUS = &H10
Private Const ODS_DEFAULT = &H20
Private Const ODS_COMBOBOXEDIT = &H1000


Private Const LF_FACESIZE = 32
Private Const SYMBOL_CHARSET = 2

Private Const LOGPIXELSY = 90

Private Const FW_NORMAL = 400
Private Const FW_BOLD = 700


Private Const GWL_WNDPROC = -4


Private Const NOTSRCERASE = &H1100A6
Private Const NOTSRCCOPY = &H330008
Private Const SRCERASE = &H440328
Private Const SRCINVERT = &H660046
Private Const SRCAND = &H8800C6
Private Const MERGEPAINT = &HBB0226
Private Const MERGECOPY = &HC000CA
Private Const SRCCOPY = &HCC0020
Private Const SRCPAINT = &HEE0086
Private Const PATPAINT = &HFB0A09

Private Const BLACKNESS = &H42
Private Const DSTINVERT = &H550009
Private Const PATINVERT = &H5A0049
Private Const PATCOPY = &HF00021
Private Const WHITENESS = &HFF0062

Private Const MAGICROP = &HB8074A


Private Const TRANSPARENT = 1
Private Const OPAQUE = 2


Private Const DT_BOTTOM = &H8
Private Const DT_CALCRECT = &H400
Private Const DT_CENTER = &H1
Private Const DT_EXPANDTABS = &H40
Private Const DT_EXTERNALLEADING = &H200
Private Const DT_INTERNAL = &H1000
Private Const DT_LEFT = &H0
Private Const DT_NOCLIP = &H100
Private Const DT_NOPREFIX = &H800
Private Const DT_RIGHT = &H2
Private Const DT_SINGLELINE = &H20
Private Const DT_TABSTOP = &H80
Private Const DT_TOP = &H0
Private Const DT_VCENTER = &H4
Private Const DT_WORDBREAK = &H10

Private Const ODT_MENU = 1

Private Const MNC_IGNORE = 0
Private Const MNC_CLOSE = 1
Private Const MNC_EXECUTE = 2
Private Const MNC_SELECT = 3


Private Const MIIM_STATE = &H1&
Private Const MIIM_ID = &H2
Private Const MIIM_SUBMENU = &H4
Private Const MIIM_CHECKMARKS = &H8
Private Const MIIM_TYPE = &H10
Private Const MIIM_DATA = &H20
Private Const MIIM_STRING = &H40
Private Const MIIM_BITMAP = &H80
Private Const MIIM_FTYPE = &H100

Private Const MF_INSERT = &H0
Private Const MF_CHANGE = &H80
Private Const MF_APPEND = &H100
Private Const MF_DELETE = &H200
Private Const MF_REMOVE = &H1000

Private Const MF_BYCOMMAND = &H0
Private Const MF_BYPOSITION = &H400

Private Const MF_SEPARATOR = &H800

Private Const MF_ENABLED = &H0
Private Const MF_GRAYED = &H1
Private Const MF_DISABLED = &H2

Private Const MF_UNCHECKED = &H0
Private Const MF_CHECKED = &H8
Private Const MF_USECHECKBITMAPS = &H200

Private Const MF_STRING = &H0
Private Const MF_BITMAP = &H4
Private Const MF_OWNERDRAW = &H100

Private Const MF_POPUP = &H10
Private Const MF_MENUBARBREAK = &H20
Private Const MF_MENUBREAK = &H40

Private Const MF_UNHILITE = &H0
Private Const MF_HILITE = &H80

Private Const MF_DEFAULT = &H1000
Private Const MF_SYSMENU = &H2000
Private Const MF_HELP = &H4000
Private Const MF_RIGHTJUSTIFY = &H4000

Private Const MF_MOUSESELECT = &H8000
Private Const MF_END = &H80

Private Const MFT_STRING = MF_STRING
Private Const MFT_BITMAP = MF_BITMAP
Private Const MFT_MENUBARBREAK = MF_MENUBARBREAK
Private Const MFT_MENUBREAK = MF_MENUBREAK
Private Const MFT_OWNERDRAW = MF_OWNERDRAW
Private Const MFT_RADIOCHECK = &H200
Private Const MFT_SEPARATOR = MF_SEPARATOR
Private Const MFT_RIGHTORDER = &H2000
Private Const MFT_RIGHTJUSTIFY = MF_RIGHTJUSTIFY

Private Const MFS_GRAYED = &H3
Private Const MFS_DISABLED = MFS_GRAYED
Private Const MFS_CHECKED = MF_CHECKED
Private Const MFS_HILITE = MF_HILITE
Private Const MFS_ENABLED = MF_ENABLED
Private Const MFS_UNCHECKED = MF_UNCHECKED
Private Const MFS_UNHILITE = MF_UNHILITE
Private Const MFS_DEFAULT = MF_DEFAULT
'Private Const MFS_MASK = &H108B
'Private Const MFS_HOTTRACKDRAWN = &H10000000
'Private Const MFS_CACHEDBMP = &H20000000
'Private Const MFS_BOTTOMGAPDROP = &H40000000
'Private Const MFS_TOPGAPDROP = &H80000000
'Private Const MFS_GAPDROP = &HC0000000


Private Const CXGAP = 1
Private Const CXTEXTMARGIN = 2
Private Const CXBUTTONMARGIN = 2
Private Const CYBUTTONMARGIN = 2


Private Const BDR_RAISEDOUTER = &H1
Private Const BDR_SUNKENOUTER = &H2
Private Const BDR_RAISEDINNER = &H4
Private Const BDR_SUNKENINNER = &H8

Private Const BDR_OUTER = &H3
Private Const BDR_INNER = &HC

Private Const EDGE_RAISED = (BDR_RAISEDOUTER Or BDR_RAISEDINNER)
Private Const EDGE_SUNKEN = (BDR_SUNKENOUTER Or BDR_SUNKENINNER)
Private Const EDGE_ETCHED = (BDR_SUNKENOUTER Or BDR_RAISEDINNER)
Private Const EDGE_BUMP = (BDR_RAISEDOUTER Or BDR_SUNKENINNER)


Private Const BF_LEFT = &H1
Private Const BF_TOP = &H2
Private Const BF_RIGHT = &H4
Private Const BF_BOTTOM = &H8

Private Const BF_TOPLEFT = (BF_TOP Or BF_LEFT)
Private Const BF_TOPRIGHT = (BF_TOP Or BF_RIGHT)
Private Const BF_BOTTOMLEFT = (BF_BOTTOM Or BF_LEFT)
Private Const BF_BOTTOMRIGHT = (BF_BOTTOM Or BF_RIGHT)
Private Const BF_RECT = (BF_LEFT Or BF_TOP Or BF_RIGHT Or BF_BOTTOM)

Private Const BF_DIAGONAL = &H10


Private Const BF_DIAGONAL_ENDTOPRIGHT = (BF_DIAGONAL Or BF_TOP Or BF_RIGHT)
Private Const BF_DIAGONAL_ENDTOPLEFT = (BF_DIAGONAL Or BF_TOP Or BF_LEFT)
Private Const BF_DIAGONAL_ENDBOTTOMLEFT = (BF_DIAGONAL Or BF_BOTTOM Or BF_LEFT)
Private Const BF_DIAGONAL_ENDBOTTOMRIGHT = (BF_DIAGONAL Or BF_BOTTOM Or BF_RIGHT)

Private Const BF_MIDDLE = &H800
Private Const BF_SOFT = &H1000
Private Const BF_ADJUST = &H2000
Private Const BF_FLAT = &H4000
Private Const BF_MONO = &H8000


Private Const WM_WINDOWPOSCHANGED = &H47
Private Const WM_SYSCOLORCHANGE = &H15
Private Const WM_NCMOUSEMOVE = &HA0
Private Const WM_COMMAND = &H111
Private Const WM_CLOSE = &H10
Private Const WM_DRAWITEM = &H2B
Private Const WM_GETFONT = &H31
Private Const WM_MEASUREITEM = &H2C
Private Const WM_NCHITTEST = &H84
Private Const WM_MENUSELECT = &H11F
Private Const WM_MENUCHAR = &H120
Private Const WM_INITMENUPOPUP = &H117
Private Const WM_ENTERMENULOOP = &H211
Private Const WM_INITMENU = &H116
Private Const WM_WININICHANGE = &H1A
Private Const WM_SETCURSOR = &H20
Private Const WM_SETTINGCHANGE = WM_WININICHANGE
Private Const WM_CANCELMODE = &H1F
Private Const WM_MDISETMENU = &H230
Private Const WM_MDIREFRESHMENU = &H234

Private Type POINTAPI
  X                             As Long
  Y                             As Long
End Type

Private Type RECT
  Left                          As Long
  Top                           As Long
  Right                         As Long
  Bottom                        As Long
End Type

Private Type OSVERSIONINFO
  dwOSVersionInfoSize           As Long
  dwMajorVersion                As Long
  dwMinorVersion                As Long
  dwBuildNumber                 As Long
  dwPlatformId                  As Long
  szCSDVersion(1 To 128)        As Byte
End Type

Private Const VER_PLATFORM_WIN32_NT = 2
Private Const VER_PLATFORM_WIN32_WINDOWS = 1
Private Const VER_PLATFORM_WIN32s = 0

Private Type MEASUREITEMSTRUCT
  CtlType                       As Long
  CtlID                         As Long
  itemID                        As Long
  itemWidth                     As Long
  itemHeight                    As Long
  ItemData                      As Long
End Type

Private Type DRAWITEMSTRUCT
  CtlType                       As Long
  CtlID                         As Long
  itemID                        As Long
  itemAction                    As Long
  itemState                     As Long
  hwndItem                      As Long
  hdc                           As Long
  rcItem                        As RECT
  ItemData                      As Long
End Type

Private Type MENUITEMINFO
  cbSize                        As Long
  fMask                         As Long
  fType                         As Long
  fState                        As Long
  wID                           As Long
  hSubMenu                      As Long
  hbmpChecked                   As Long
  hbmpUnchecked                 As Long
  dwItemData                    As Long
  dwTypeData                    As Long
  cch                           As Long
End Type

Private Type LogFont
  lfHeight                      As Long
  lfWidth                       As Long
  lfEscapement                  As Long
  lfOrientation                 As Long
  lfWeight                      As Long
  lfItalic                      As Byte
  lfUnderline                   As Byte
  lfStrikeOut                   As Byte
  lfCharSet                     As Byte
  lfOutPrecision                As Byte
  lfClipPrecision               As Byte
  lfQuality                     As Byte
  lfPitchAndFamily              As Byte
  lfFaceName                    As String * 32
End Type

Private Type TEXTMETRIC
  tmHeight                      As Long
  tmAscent                      As Long
  tmDescent                     As Long
  tmInternalLeading             As Long
  tmExternalLeading             As Long
  tmAveCharWidth                As Long
  tmMaxCharWidth                As Long
  tmWeight                      As Long
  tmOverhang                    As Long
  tmDigitizedAspectX            As Long
  tmDigitizedAspectY            As Long
  tmFirstChar                   As Byte
  tmLastChar                    As Byte
  tmDefaultChar                 As Byte
  tmBreakChar                   As Byte
  tmItalic                      As Byte
  tmUnderlined                  As Byte
  tmStruckOut                   As Byte
  tmPitchAndFamily              As Byte
  tmCharSet                     As Byte
End Type

Private Type IMAGEINFO
  hbmImage                      As Long
  hbmMask                       As Long
  Unused1                       As Long
  Unused2                       As Long
  rcImage                       As RECT
End Type

Private Type ButType
    ButText                     As String
    ButImage                    As Integer
End Type

Private m_bmpChecked            As Long
Private m_bmpRadioed            As Long

Private m_MarlettFont           As Long
Private m_iBitmapWidth          As Integer
Private m_SideBitmapWidth       As Long

Private pmds                    As clsMyItemDatas
Private WndCol                  As Collection

Private Sub ConvertMenu(hwnd As Long, hMenu As Long, nIndex As Long, bSysMenu As Boolean, bShowButtons As Boolean, Optional Permanent As Boolean = False)
    On Error GoTo ErrorHandle
    Dim i                           As Long
    Dim k                           As Byte
    Dim info                        As MENUITEMINFO
    Dim TmpMenuInfo                 As ButType
    Dim dwItemData                  As Long
    Dim pmd                         As clsMyItemData
    Dim Text                        As String
    Dim ByteBuffer()                As Byte
    Dim nItem                       As Long
    nItem& = GetMenuItemCount(hMenu&)
    If nItem& = -1 Then Exit Sub
    For i& = 0 To nItem& - 1
        ReDim ByteBuffer(0 To 200) As Byte
        For k = 0 To 200
            ByteBuffer(k) = 0
        Next k
        info.fMask = MIIM_DATA Or MIIM_ID Or MIIM_TYPE Or MIIM_SUBMENU
        info.dwTypeData = VarPtr(ByteBuffer(0))
        info.cch = UBound(ByteBuffer)
        info.cbSize = LenB(info)
        Call GetMenuItemInfo(hMenu&, i&, MF_BYPOSITION, info)
        dwItemData& = info.dwItemData
        If bSysMenu And (info.wID >= &HF000) Then GoTo NextGoto
        info.fMask = 0& 'reset mask value
        If bShowButtons Then
            If Not CBool(info.fType And MFT_OWNERDRAW) Then
                info.fType = info.fType Or MFT_OWNERDRAW
                info.fMask = info.fMask Or MIIM_TYPE
                If dwItemData& = 0& Then
                    info.dwItemData = CLng(pmds.Count + 1)
                    info.fMask = info.fMask Or MIIM_DATA
                    Set pmd = pmds.Add(CStr(info.dwItemData))
                    Text$ = Left(StrConv(ByteBuffer, vbUnicode), info.cch)
                    pmd.sMenuText = Text$
                    Dim iBreakPos As Integer
                    iBreakPos% = InStr(Text$, "|")
                    If iBreakPos% Then
                        Dim iBreak2Pos As Integer
                        iBreak2Pos% = InStr(Right(Text$, Len(Text$) - iBreakPos%), "|")
                        Dim HelpText As String
                        Dim iHelpLen As Integer
                        HelpText$ = Mid(Text$, iBreakPos% + 1, iBreak2Pos% - 1)
                        iHelpLen% = Len(HelpText$)
                        pmd.sMenuHelp = HelpText$
                        pmd.sMenuText = Right(Text$, Len(Text$) - (iBreakPos% + iBreak2Pos%))
                    Else
                        pmd.sMenuText = Text$
                    End If
                    Dim cFirstChar As String * 3
                    cFirstChar$ = Left(Text$, 3)
                    Dim iCount As Boolean
                    iCount = False
                    If cFirstChar$ = "(-)" Then
                        iCount = True
                        info.fType = info.fType Or &H800
                        If pmd.sMenuHelp = "" Then pmd.sMenuText = Right(Text$, Len(Text$) - 3)
                    ElseIf cFirstChar$ = "(=)" Then
                        info.fType = info.fType Or &H40
                        If pmd.sMenuHelp = "" Then pmd.sMenuText = Right(Text$, Len(Text$) - 3)
                    ElseIf cFirstChar$ = "(+)" Then
                        info.fType = info.fType Or &H40 Or &H800
                        If pmd.sMenuHelp = "" Then pmd.sMenuText = Right(Text$, Len(Text$) - 3)
                    End If
                    pmd.bAsMark = (cFirstChar$ = "(*)") Or (cFirstChar$ = "(#)")
                    If pmd.bAsMark Then
                        pmd.bAsCheck = (cFirstChar$ = "(#)")
                        If pmd.sMenuHelp = "" Then pmd.sMenuText = Right(Text$, Len(Text$) - 3)
                    End If
                    TmpMenuInfo = GetButtonIndex(hwnd&, pmd.sMenuText)
                    pmd.iButton = TmpMenuInfo.ButImage
                    pmd.sMenuText = TmpMenuInfo.ButText
                    pmd.fType = info.fType
                    pmd.bTrueSub = (info.hSubMenu <> 0&) And (Not Permanent)
                Else
                    Set pmd = pmds.Item(CStr(dwItemData&))
                End If
                pmd.bMainMenu = Permanent
            End If
            If Not Permanent Then Call WndCol(CStr(hwnd&)).AddMenuHead(hMenu)
        Else
            If info.fType And MFT_OWNERDRAW Then
                info.fType = info.fType And (Not MFT_OWNERDRAW)
                info.fMask = info.fMask Or MIIM_TYPE
                Set pmd = pmds.Item(CStr(dwItemData&))
                Dim cLeadChar As String
                cLeadChar$ = ""
                If pmd.bAsMark Then
                If pmd.bAsCheck Then
                    cLeadChar = "(#)"
                Else
                    cLeadChar = "(*)"
                End If
            End If
            If pmd.fType And &H800 Then
                cLeadChar$ = "(-)"
                info.fType = info.fType And (Not &H800)
            End If
            If pmd.fType And &H40 Then
                cLeadChar$ = "(=)"
                info.fType = info.fType And (Not &H40)
            End If
            If pmd.sMenuHelp <> "" Then _
                pmd.sMenuHelp = "|" + pmd.sMenuHelp + "|"
                Text$ = cLeadChar$ + pmd.sMenuHelp + pmd.sMenuText
                info.cch = BSTRtoLPSTR(Text$, ByteBuffer, info.dwTypeData)
            End If
            If dwItemData <> 0& Then
                info.dwItemData = 0&
                info.fMask = info.fMask Or MIIM_DATA
                pmds.Remove CStr(dwItemData&) 'by key
            End If
        End If
        If info.fMask Then Call SetMenuItemInfo(hMenu&, i&, MF_BYPOSITION, info)
NextGoto:
    Next i&
    Exit Sub
ErrorHandle:
  Debug.Print Err.Number; Err.Description; " ConvertMenu"
  Err.Clear
End Sub
Private Sub OnInitMenuPopup(hwnd As Long, hMenu As Long, nIndex As Long, bSysMenu As Boolean)
    WndCol(CStr(hwnd&)).MainPopedIndex = -2 ' Deselect main menu item
    Call ConvertMenu(hwnd&, hMenu&, nIndex&, bSysMenu, True, False)
End Sub
Private Function OnMenuChar(nChar As Long, nFlags As Long, hMenu As Long) As Long
    Dim i                       As Long
    Dim nItem                   As Long
    Dim dwItemData              As Long
    Dim info                    As MENUITEMINFO
    Dim Count                   As Integer: Count% = 0
    Dim iCurrent                As Integer
    Dim Text                    As String
    Dim iAmpersand              As Integer
    Dim bMainMenu               As Boolean
    Dim iSelect                 As Integer
    
    ReDim ItemMatch(0 To 0) As Integer
    nItem& = GetMenuItemCount(hMenu&)
    For i& = 0 To nItem& - 1
        info.cbSize = LenB(info)
        info.fMask = MIIM_DATA Or MIIM_TYPE Or MIIM_STATE
        Call GetMenuItemInfo(hMenu&, i&, MF_BYPOSITION, info)
        dwItemData& = info.dwItemData
        If (info.fType And MFT_OWNERDRAW) And dwItemData& <> 0 Then
            Text$ = pmds(CStr(dwItemData&)).sMenuText
            iAmpersand% = InStr(Text$, "&")
            If (iAmpersand% > 0) And (UCase(Chr(nChar&)) = UCase(Mid(Text$, iAmpersand% + 1, 1))) Then
                If Count > UBound(ItemMatch) Then ReDim Preserve ItemMatch(0 To Count%)
                ItemMatch(Count%) = i&
                Count% = Count% + 1
            End If
        End If
        If info.fState And MFS_HILITE Then iCurrent% = i&
    Next i&
    Count% = Count% - 1
    If Count% = -1 Then
        OnMenuChar = 0&
        Exit Function
    End If
    bMainMenu = pmds(CStr(dwItemData&)).bMainMenu
    If Count% = 0 Then
        OnMenuChar = MakeLong(ItemMatch(0), MNC_EXECUTE)
        Exit Function
    End If
    For i& = 0 To Count%
        If ItemMatch(i&) = iCurrent% Then
            iSelect% = i&
            Exit For
        End If
    Next i&
    OnMenuChar = MakeLong(ItemMatch(iSelect%), MNC_SELECT)
End Function
Private Sub DrawMenuText(hwnd As Long, hdc As Long, rc As RECT, Text As String, Color As Long, Optional bLeftAlign As Boolean = True, Optional bRightToLeft As Boolean = False)
    Dim LeftStr                 As String
    Dim RightStr                As String
    Dim iTabPos                 As Integer
    Dim OldFont                 As Long
    LeftStr$ = Text$
    iTabPos = InStr(LeftStr$, Chr(9))
    If iTabPos > 0 Then
        RightStr$ = Right$(LeftStr$, Len(LeftStr$) - iTabPos)
        LeftStr$ = Left$(LeftStr$, iTabPos - 1)
    End If
    Call SetTextColor(hdc&, Color&)
    OldFont& = SelectObject(hdc&, GetMenuFont(hwnd&))
    Call DrawText(hdc&, LeftStr$, Len(LeftStr$), rc, IIf(bLeftAlign, IIf(bRightToLeft, DT_RIGHT, DT_LEFT), DT_CENTER) Or DT_VCENTER Or DT_SINGLELINE)
    If iTabPos > 0 Then Call DrawText(hdc&, RightStr$, Len(RightStr$), rc, IIf(bRightToLeft, DT_LEFT, DT_RIGHT) Or DT_VCENTER Or DT_SINGLELINE)
    Call SelectObject(hdc&, OldFont&)
End Sub
Private Function OnDrawItem(hwnd As Long, ByRef dsPtr As Long, Optional bOverMain As Boolean = False) As Boolean
    On Error GoTo ErrHandler
    Dim lpds                    As DRAWITEMSTRUCT
    Call CopyMemory(lpds, ByVal dsPtr&, Len(lpds))
    Dim rt                      As RECT
    Dim rtItem                  As RECT
    Dim rtText                  As RECT
    Dim rtButn                  As RECT
    Dim rtIcon                  As RECT
    Dim rtHighlight             As RECT
    Dim OldFont                 As Long
    Dim hIcon                   As Long
    Dim pic                     As StdPicture
    Dim dwItemData              As Long
    Dim hdc                     As Long
    Dim WndObj                  As clsWndCoolMenu: Set WndObj = WndCol(CStr(hwnd&))
    Dim pmd                     As clsMyItemData
    Dim SepMargin               As Integer
    Dim bDisabled               As Boolean
    Dim bSelected               As Boolean
    Dim bChecked                As Boolean
    Dim bHaveButn               As Boolean
    Dim iButton                 As Integer
    Dim info                    As MENUITEMINFO
    Dim iButnWidth              As Integer
    Dim dwColorBG               As Long
    Dim dwSelTextColor          As Long
    Dim dwColorText             As Long
    Dim TextOffset              As Integer
    Dim rtArrow                 As RECT
    dwItemData& = lpds.ItemData
    If (dwItemData& = 0&) Or (lpds.CtlType <> ODT_MENU) Or (dwItemData& > pmds.Count) Then
        OnDrawItem = False
        Exit Function
    End If
    hdc& = lpds.hdc
    LSet rtItem = lpds.rcItem
    Set pmd = pmds.Item(CStr(dwItemData&))
    
    If pmd.fType And MFT_SEPARATOR Then
        LSet rt = rtItem
        LSet rtText = rtItem
        SepMargin = 15
        rt.Left = rt.Left + SepMargin
        rt.Right = rt.Right - SepMargin
        rt.Top = rt.Top + ((rt.Bottom - rt.Top) \ 2) - 1
        Call DrawEdge(hdc&, rt, EDGE_ETCHED, BF_TOP)
        If pmd.sMenuText <> "" Then
            OldFont& = SelectObject(hdc&, GetMenuFontSep(hwnd&))
            rtText = OffsetRect(rtText, 1, 1)
            Call SetBkMode(hdc&, OPAQUE)
            Call SetTextColor(hdc&, GetSysColor(COLOR_BTNLIGHT))
            Call DrawText(hdc&, " " + pmd.sMenuText + " ", 2 + Len(pmd.sMenuText), rtText, DT_CENTER Or DT_VCENTER Or DT_SINGLELINE)
            rtText = OffsetRect(rtText, -1, -1)
            Call SetBkMode(hdc&, TRANSPARENT)
            Call SetTextColor(hdc&, GetSysColor(COLOR_BTNSHADOW))
            Call DrawText(hdc&, " " + pmd.sMenuText + " ", 2 + Len(pmd.sMenuText), rtText, DT_CENTER Or DT_VCENTER Or DT_SINGLELINE)
            Call SelectObject(hdc&, OldFont&)
        End If
    ElseIf Left(pmd.sMenuText, 1) = "!" Then
        Dim SideBitmap As Long
        Dim sBitmap As String: sBitmap = "c:\win95\bureau\smart.bmp" + Chr(0)
        Dim hmemDC As Long
        Dim hOldBitmap As Long
    Else
        bDisabled = lpds.itemState And ODS_GRAYED
        bSelected = lpds.itemState And ODS_SELECTED
        bChecked = lpds.itemState And ODS_CHECKED
        bHaveButn = False
        iButton = pmd.iButton
        LSet rtButn = rtItem
        If WndObj.RightToLeft Then
            rtButn.Left = rtButn.Right - (m_iBitmapWidth + CXBUTTONMARGIN)
        Else
            rtButn.Right = rtButn.Left + m_iBitmapWidth + CXBUTTONMARGIN
        End If
        If iButton >= 0 Then
            bHaveButn = True
            rtIcon.Left = rtButn.Left + (CXBUTTONMARGIN \ 2)
            rtIcon.Right = rtIcon.Left + m_iBitmapWidth
            rtIcon.Top = rtButn.Top + ((rtButn.Bottom - rtButn.Top) - m_iBitmapWidth) \ 2
            rtIcon.Bottom = rtIcon.Top + m_iBitmapWidth
            If Not bDisabled Then
                Call FillRectEx(hdc&, rtButn, GetSysColor(IIf(bChecked And (Not bSelected), COLOR_BTNLIGHT, COLOR_MENU)))
                If bSelected Or bChecked Then Call DrawEdge(hdc&, rtButn, IIf(bChecked, BDR_SUNKENOUTER, BDR_RAISEDINNER), BF_RECT)
                Set pic = ConvertTo16(LoadResPicture(PrepTextForImage(pmd.sMenuText), vbResIcon).Handle)
                hIcon = pic.Handle
                Call DrawState(hdc&, 0&, 0&, hIcon&, 0&, rtIcon.Left, rtIcon.Top, rtIcon.Left, rtIcon.Top, DST_ICON Or DSS_NORMAL)
            Else
                Set pic = ConvertTo16(LoadResPicture(PrepTextForImage(pmd.sMenuText), vbResIcon).Handle)
                hIcon = pic.Handle
                Call DrawState(hdc&, 0&, 0&, hIcon&, 0&, rtIcon.Left, rtIcon.Top, rtIcon.Left + m_iBitmapWidth%, rtIcon.Top + m_iBitmapWidth%, DST_ICON Or DSS_DISABLED)
            End If
        Else
            info.cbSize = LenB(info)
            info.fMask = MIIM_CHECKMARKS
            Call GetMenuItemInfo(lpds.hwndItem, lpds.itemID, MF_BYCOMMAND, info)
            If bChecked Or CBool(info.hbmpUnchecked) Or (pmd.bAsMark And WndObj.ComplexChecks) Then
                bHaveButn = Draw3DMark(hwnd&, hdc&, rtButn, bChecked, bSelected, bDisabled, IIf(bChecked, info.hbmpChecked, info.hbmpUnchecked), pmd.bAsCheck)
            End If
        End If
        iButnWidth% = m_iBitmapWidth% + CXBUTTONMARGIN
        dwColorBG = IIf(bSelected And WndObj.FullSelect, WndObj.SelectColor&, GetSysColor(COLOR_MENU))
        LSet rtText = rtItem
        If pmd.bMainMenu Then Call FillRectEx(hdc&, rtItem, GetSysColor(COLOR_MENU))
        If (bSelected Or (lpds.itemAction = ODA_SELECT)) And (Not bDisabled) Then
            LSet rtHighlight = rtItem
            If bHaveButn Then
                If WndObj.RightToLeft Then
                    rtHighlight.Right = rtItem.Right - (iButnWidth% + CXGAP)
                Else
                    rtHighlight.Left = rtItem.Left + iButnWidth% + CXGAP
                End If
            End If
            If pmd.bMainMenu And bSelected Then
                rtText = OffsetRect(rtText, 2, 1)
                Call DrawEdge(hdc&, rtHighlight, BDR_SUNKENOUTER, BF_RECT)
            Else
                Call FillRectEx(hdc&, rtHighlight, dwColorBG&)
            End If
        End If
        If Not pmd.bMainMenu Then
            If WndObj.RightToLeft Then
                rtText.Right = rtItem.Right - (iButnWidth% + CXGAP + CXTEXTMARGIN)
                rtText.Left = rtItem.Left + iButnWidth%
            Else
                rtText.Left = rtItem.Left + iButnWidth% + CXGAP + CXTEXTMARGIN
                rtText.Right = rtItem.Right - iButnWidth%
            End If
        End If
        Call SetBkMode(hdc&, TRANSPARENT)
        dwSelTextColor& = GetSysColor(COLOR_HIGHLIGHTTEXT)
        dwColorText& = IIf(bDisabled, GetSysColor(COLOR_GRAYTEXT), IIf(bSelected And (Not pmd.bMainMenu), IIf(WndObj.FullSelect, dwSelTextColor&, WndObj.SelectColor&), IIf(WndObj.ForeColor& = 0&, GetSysColor(COLOR_MENUTEXT), WndObj.ForeColor&)))
        TextOffset = 1
        If bDisabled Then Call DrawMenuText(hwnd&, hdc&, OffsetRect(rtText, TextOffset, TextOffset), pmds(CStr(dwItemData)).sMenuText, GetSysColor(COLOR_BTNHIGHLIGHT), Not pmd.bMainMenu, WndObj.RightToLeft)
        Call DrawMenuText(hwnd&, hdc&, rtText, pmd.sMenuText, dwColorText&, Not pmd.bMainMenu, WndObj.RightToLeft)
    End If
    If pmd.bTrueSub Then
        LSet rtArrow = rtItem
        If WndObj.RightToLeft Then
            rtArrow.Left = rtArrow.Left + CXTEXTMARGIN
        Else
            rtArrow.Right = rtArrow.Right - CXTEXTMARGIN
        End If
        rtArrow.Top = rtArrow.Top + CXTEXTMARGIN
        Call PrintGlyph(hdc&, IIf(WndObj.RightToLeft, "3", "4"), dwColorText&, rtArrow, IIf(WndObj.RightToLeft, DT_LEFT, DT_RIGHT) Or DT_TOP Or DT_SINGLELINE)
        Call ExcludeClipRect(hdc&, rtItem.Left, rtItem.Top, rtItem.Right, rtItem.Bottom)
    End If
    Call CopyMemory(ByVal dsPtr&, lpds, Len(lpds))
    Set WndObj = Nothing
    OnDrawItem = True
    Exit Function
ErrHandler:
    Debug.Print Err.Number; Err.Description; " OnDrawItem"
    Err.Clear
End Function
Private Function OnMeasureItem(hwnd As Long, ByRef miPtr As Long) As Boolean
    Dim lpms                    As MEASUREITEMSTRUCT
    Dim dwItemData              As Long
    Dim rc                      As RECT
    Dim rcHeight                As Integer
    Dim OldFont                 As Long
    Dim hWndDC                  As Long
    Dim pmd                     As clsMyItemData
    Dim iCYMENU                 As Integer
    Dim itemWidth               As Long
    Call CopyMemory(lpms, ByVal miPtr, Len(lpms))
    dwItemData& = lpms.ItemData
    If (dwItemData& = 0&) Or (lpms.CtlType <> ODT_MENU) Then
        OnMeasureItem = False
        Exit Function
    End If
    Set pmd = pmds.Item(CStr(dwItemData&))
    iCYMENU% = GetSystemMetrics(SM_CYMENU)
    If pmd.fType And MFT_SEPARATOR Then
        hWndDC& = GetDC(hwnd&)
        OldFont& = SelectObject(hWndDC&, GetMenuFont(hwnd&))
        rcHeight = DrawText(hWndDC&, "A", 1&, rc, DT_SINGLELINE Or DT_CALCRECT) + 1
        lpms.itemHeight = IIf(iCYMENU% \ 2 > rcHeight, iCYMENU% \ 2, rcHeight)
        lpms.itemWidth = 0
        Call SelectObject(hWndDC&, OldFont&)
        Call ReleaseDC(hwnd&, hWndDC&)
    ElseIf Left(pmd.sMenuText, 1) = "!" Then
        lpms.itemHeight = 0
        lpms.itemWidth = 0
    Else
        hWndDC& = GetDC(hwnd&)
        OldFont& = SelectObject(hWndDC&, GetMenuFont(hwnd&))
        Call DrawText(hWndDC&, pmd.sMenuText, Len(pmd.sMenuText), rc, DT_LEFT Or DT_SINGLELINE Or DT_CALCRECT) 'Or DT_VCENTER
        Call SelectObject(hWndDC&, OldFont&)
        Call ReleaseDC(hwnd&, hWndDC&)
        rcHeight = rc.Bottom - rc.Top
        lpms.itemHeight = IIf(rcHeight > iCYMENU%, rcHeight, iCYMENU%)
        itemWidth& = (rc.Right - rc.Left)
        If Not pmd.bMainMenu Then
            itemWidth& = itemWidth& + (CXTEXTMARGIN * 2) + CXGAP + (m_iBitmapWidth% + CXBUTTONMARGIN) * 2
            itemWidth& = itemWidth& - (GetSystemMetrics(SM_CXMENUCHECK) - 1)
        End If
        lpms.itemWidth = itemWidth& + m_SideBitmapWidth
    End If
    Call CopyMemory(ByVal miPtr, lpms, Len(lpms))
    OnMeasureItem = True
End Function
Public Function GetMenuFont(hwnd As Long, Optional bForceReset As Boolean = False) As Long
    Dim WndObj                  As clsWndCoolMenu
    Dim sText                   As String
    Dim TextLen                 As Long
    Dim tLF                     As LogFont
    Dim tm                      As TEXTMETRIC
    Dim hWndDC                  As Long: hWndDC& = GetDC(hwnd&)
    Dim hdc                     As Long: hdc& = GetWindowDC(hwnd&)
    Set WndObj = WndCol(CStr(hwnd&))
    If (WndObj.MenuFont = 0) Or bForceReset Then
        If WndObj.FontName = "" Then
            sText$ = Space$(255)
            TextLen& = Len(sText$)
            TextLen& = GetTextFace(hWndDC&, TextLen&, sText$)
            WndObj.FontName = Left$(sText$, TextLen&)
            If WndObj.ForeColor = 0& Then WndObj.ForeColor = GetTextColor(hWndDC&)
            Call GetTextMetrics(hWndDC&, tm)
            Call ReleaseDC(hwnd&, hWndDC&)
            tLF.lfHeight = tm.tmHeight
            tLF.lfWeight = tm.tmWeight
        Else
            tLF.lfWeight = FW_NORMAL
            tLF.lfHeight = -MulDiv(WndObj.FontSize&, GetDeviceCaps(hdc&, LOGPIXELSY), 72)
            Call ReleaseDC(hwnd&, hdc&)
        End If
        tLF.lfFaceName = WndObj.FontName$ + Chr(0)
        WndObj.MenuFont& = CreateFontIndirect(tLF)
    End If
    GetMenuFont& = WndObj.MenuFont&
    Set WndObj = Nothing
End Function
Private Function GetMenuFontSep(hwnd As Long) As Long
    Dim WndObj                  As clsWndCoolMenu
    Dim tLF                     As LogFont
    Set WndObj = WndCol(CStr(hwnd&))
    If WndObj.MenuFontSep& = 0& Then
        tLF.lfFaceName = "Small Fonts" + Chr(0)
        tLF.lfHeight = 11
        tLF.lfWeight = FW_NORMAL
        WndObj.MenuFontSep& = CreateFontIndirect(tLF)
    End If
    GetMenuFontSep& = WndObj.MenuFontSep&
    Set WndObj = Nothing
End Function
Public Function Install(wndHandle As Long, Optional HelpObj As clsHelpCallBack) As Boolean
    Dim NewWnd                      As clsWndCoolMenu
    m_iBitmapWidth% = 16
    m_SideBitmapWidth = 0
    If wndHandle <> 0 Then
        If WndCol Is Nothing Then
            Set WndCol = New Collection
            Set pmds = New clsMyItemDatas
        End If
        Set NewWnd = New clsWndCoolMenu
        NewWnd.hwnd = wndHandle&
        NewWnd.PrevProc = GetWindowLong(wndHandle&, GWL_WNDPROC)
        NewWnd.SelectColor = GetSysColor(COLOR_HIGHLIGHT)
        Call SetWindowLong(wndHandle&, GWL_WNDPROC, AddressOf WindowProc)
        If Not (HelpObj Is Nothing) Then Set NewWnd.HelpObj = HelpObj
        NewWnd.SCMainMenu = True
        WndCol.Add NewWnd, CStr(wndHandle&)
        Set NewWnd = Nothing
        Call ConvertMenu(wndHandle&, GetMenu(wndHandle&), 0&, False, True, True)
    End If
    Install = True
End Function
Public Function Uninstall(wndHandle As Long) As Boolean
    If (wndHandle <> 0) And (Not (WndCol Is Nothing)) Then
        Call SetWindowLong(wndHandle&, GWL_WNDPROC, WndCol(CStr(wndHandle&)).PrevProc)
        WndCol.Remove CStr(wndHandle&)
        If WndCol.Count = 0 Then
            Set WndCol = Nothing
            Call DeleteObject(m_MarlettFont&)
            Call DeleteObject(m_bmpChecked&)
            Call DeleteObject(m_bmpRadioed)
            Set pmds = Nothing
        End If
        Uninstall = True
    End If
End Function
Private Sub FillRectEx(hdc As Long, rc As RECT, Color As Long)
    Dim hOldBrush               As Long
    Dim hNewBrush               As Long
    hNewBrush& = CreateSolidBrush(Color&)
    Call FillRect(hdc&, rc, hNewBrush&)
    Call DeleteObject(hNewBrush&)
End Sub
Private Function OffsetRect(InRect As RECT, ByVal xOffset As Long, ByVal yOffset As Long) As RECT
    OffsetRect.Left = InRect.Left + xOffset&
    OffsetRect.Right = InRect.Right + xOffset&
    OffsetRect.Top = InRect.Top + yOffset&
    OffsetRect.Bottom = InRect.Bottom + yOffset&
End Function
Private Sub OnMenuSelect(hwnd As Long, nItemID As Integer, nFlags As Integer, hSysMenu As Long)
    On Error GoTo ErrHandler
    Dim WndObj                  As clsWndCoolMenu: Set WndObj = WndCol(CStr(hwnd&))
    Dim info                    As MENUITEMINFO
    Dim i                       As Integer
    info.cbSize = LenB(info)
    info.fMask = MIIM_DATA Or MIIM_STATE Or MIIM_TYPE Or MIIM_ID
    Call GetMenuItemInfo(GetMenu(hwnd&), nItemID, MF_BYCOMMAND, info)
    If Not (WndObj.HelpObj Is Nothing) Then
        If (info.dwItemData <> 0&) And Not CBool(nFlags And MF_POPUP) Then
            Call WndObj.HelpObj.RaiseHelpEvent(pmds.Item(CStr(info.dwItemData)).sMenuText, pmds.Item(CStr(info.dwItemData)).sMenuHelp, Not CBool(info.fState And MFS_DISABLED))
        Else
            Call WndObj.HelpObj.RaiseHelpEvent("", "", True)
        End If
    End If
    If (hSysMenu = 0&) And (nFlags = &HFFFF) Then
        For i% = 0 To WndObj.CountMenuHeads
            Call ConvertMenu(hwnd&, WndObj.GetMenuHead(i%), 0&, False, False)
        Next
        WndObj.MainPopedIndex = -1
    End If
    Exit Sub
ErrHandler:
    Debug.Print Err.Number; Err.Description; " OnMenuSelect"
    Err.Clear
End Sub
Private Function CheckImage(Text As String) As Boolean
    On Error GoTo ErrClear
    Dim IPic                    As StdPicture
    Set IPic = LoadResPicture(PrepTextForImage(Text), vbResIcon)
    CheckImage = True
    Exit Function
ErrClear:
    Err.Clear
    CheckImage = False
End Function
Private Function PrepTextForImage(Text As String) As String
    Dim StText As String
    StText = Text
    If InStr(1, Text, " ", vbTextCompare) > 0 Then StText = Replace(StText, " ", "+")
    If InStr(1, Text, "@", vbTextCompare) > 0 Then StText = Replace(StText, "@", "AT")
    PrepTextForImage = UCase(StText)
End Function
Private Function GetButtonIndex(hwnd As Long, sMenuText As String) As ButType
    If CheckImage(sMenuText) = True Then
        GetButtonIndex.ButImage = 10
    Else
        GetButtonIndex.ButImage = -1
    End If
    GetButtonIndex.ButText = sMenuText
End Function
Private Function BSTRtoLPSTR(sBSTR As String, B() As Byte, lpsz As Long) As Long
    Dim cBytes                  As Long
    Dim sABSTR                  As String
    cBytes = LenB(sBSTR)
    ReDim B(1 To cBytes + 2) As Byte
    sABSTR = StrConv(sBSTR, vbFromUnicode)
    lpsz = StrPtr(sABSTR)
    CopyMemory B(1), ByVal lpsz, cBytes + 2
    lpsz = VarPtr(B(1))
    BSTRtoLPSTR = cBytes
End Function
Private Sub DrawEmbossed(hdc As Long, pic As StdPicture, iButnIndex As Integer, rt As RECT, bInColor As Boolean)
    Dim rcImage                 As RECT
    Dim Mypic                   As StdPicture: Set Mypic = pic
    Dim cx                      As Integer
    Dim cy                      As Integer
    Dim hmemDC                  As Long
    Dim hBitmap                 As Long
    Dim hOldBitmap              As Long
    Dim hOldBackColor           As Long
    cx = Mypic.Width
    cy = Mypic.Height
    hmemDC& = CreateCompatibleDC(hdc&)
    If bInColor Then
        hBitmap& = CreateCompatibleBitmap(hdc&, cx%, cy%)
    Else
        hBitmap& = CreateBitmap(cx%, cy%, 1, 1, vbNull)
    End If
    hOldBitmap = SelectObject(hmemDC&, hBitmap&)
    Call PatBlt(hmemDC&, 0, 0, cx%, cy%, WHITENESS)
    Call DrawState(hmemDC&, 0&, 0&, Mypic.Handle, 0&, rt.Left, rt.Top, rt.Left + m_iBitmapWidth%, rt.Top + m_iBitmapWidth%, DST_ICON Or DSS_NORMAL)
    hOldBackColor& = SetBkColor(hdc&, RGB(255, 255, 255))
    Dim hbrShadow As Long, hbrHilite As Long
    hbrShadow& = CreateSolidBrush(GetSysColor(COLOR_BTNSHADOW))
    hbrHilite& = CreateSolidBrush(GetSysColor(COLOR_BTNHIGHLIGHT))
    Dim hOldBrush As Long
    hOldBrush& = SelectObject(hdc&, hbrHilite&)
    Call BitBlt(hdc&, rt.Left + 1, rt.Top + 1, cx%, cy%, hmemDC&, 0, 0, MAGICROP)
    Call SelectObject(hdc&, hbrShadow&)
    Call BitBlt(hdc&, rt.Left, rt.Top, cx%, cy%, hmemDC&, 0, 0, MAGICROP)
    Call SelectObject(hdc&, hOldBrush&)
    Call SetBkColor(hdc&, hOldBackColor&)
    Call SelectObject(hmemDC&, hOldBitmap&)
    Call DeleteObject(hOldBrush&)
    Call DeleteObject(hbrHilite&)
    Call DeleteObject(hbrShadow&)
    Call DeleteObject(hOldBackColor&)
    Call DeleteObject(hOldBitmap&)
    Call DeleteObject(hBitmap&)
    Call DeleteDC(hmemDC&)
End Sub

Private Function Draw3DMark(hwnd As Long, hdc As Long, rc As RECT, bCheck As Boolean, bSelected As Boolean, bDisabled As Boolean, hBmp As Long, bDrawCheck As Boolean) As Boolean
    On Error GoTo hError
    Dim WndObj                  As clsWndCoolMenu: Set WndObj = WndCol(CStr(hwnd&))
    Dim cx                      As Integer
    Dim cy                      As Integer
    Dim hmemDC                  As Long
    Dim hBmpTemp                As Long
    Dim hOldBmp                 As Long
    Dim hBrush                  As Long
    Dim lbInfo                  As LogBrush
    Dim hOldBrush               As Long
    Dim rcHighLigth             As RECT
    Dim X                       As Long: X = 0
    Dim Y                       As Long: Y = 0
    Dim hOldBackColor           As Long
    Dim i                       As Integer
    Dim BitArray(0 To 3)        As Long
    Dim hPat                    As Long
    Dim hPatBrush               As Long
    cx% = rc.Right - rc.Left
    cy% = rc.Bottom - rc.Top
    If Not CBool(hBmp) Then
        If WndObj.ComplexChecks Then
            hmemDC& = CreateCompatibleDC(hdc&)
            LSet rcHighLigth = rc
            rcHighLigth.Right = rcHighLigth.Right + 1
            rcHighLigth.Left = rcHighLigth.Left - 1
            Call FillRectEx(hdc&, rcHighLigth, IIf(bSelected And (Not bDisabled) And WndObj.FullSelect, WndObj.SelectColor&, GetSysColor(COLOR_MENU)))
            If m_bmpChecked = 0& Then
                m_bmpChecked& = LoadImage(0&, CLng(OBM_CHECKBOXES), IMAGE_BITMAP, 0&, 0&, LR_DEFAULTCOLOR)
                m_bmpRadioed& = LoadImage(0&, CLng(OBM_BTNCORNERS), IMAGE_BITMAP, 0&, 0&, LR_MONOCHROME)
            End If
            lbInfo.lbStyle = BS_HOLLOW
            hBrush& = CreateBrushIndirect(lbInfo)
            hOldBrush& = SelectObject(hdc&, hBrush&)
            If bCheck Then X = X + 13
            If bDisabled Then X = X + 26
            Y = 0
            hOldBackColor& = SetBkColor(hdc&, RGB(255, 255, 255))
            If bDrawCheck Then
                hOldBmp& = SelectObject(hmemDC&, m_bmpChecked&)
                Call BitBlt(hdc&, rc.Left + 3, rc.Top + 3, 13&, 13&, hmemDC&, X&, Y&, SRCCOPY)
            Else
            Y = 13
                hOldBmp& = SelectObject(hmemDC&, m_bmpRadioed&)
                Call BitBlt(hdc&, rc.Left + 3, rc.Top + 3, 13&, 13&, hmemDC&, 0&, 0&, MERGEPAINT)
                Call SelectObject(hmemDC&, m_bmpChecked&)
                Call BitBlt(hdc&, rc.Left + 3, rc.Top + 3, 13&, 13&, hmemDC&, X&, Y&, SRCAND)
            End If
            Call SetBkColor(hdc&, hOldBackColor&)
            Call SelectObject(hmemDC&, hOldBmp&)
            Call DeleteObject(hOldBmp&)
            Call SelectObject(hdc&, hOldBrush)
            Call DeleteObject(hBrush&)
            Call DeleteDC(hmemDC&)
        Else
            If bSelected Then
                Call FillRectEx(hdc&, rc, GetSysColor(COLOR_MENU))
            Else
                For i = 0 To 3
                    BitArray(i) = MakeLong(170, 85)
                Next
                hPat& = CreateBitmap(8&, 8&, 1&, 1&, BitArray(0))
                hPatBrush& = CreatePatternBrush(hPat&)
                hOldBrush& = SelectObject(hdc&, hPatBrush&)
                Call SetBkColor(hdc&, GetSysColor(COLOR_MENU))
                Call SetTextColor(hdc&, GetSysColor(COLOR_BTNHIGHLIGHT))
                Call PatBlt(hdc&, rc.Left, rc.Top, cx%, cy%, PATCOPY)
                Call SelectObject(hdc&, hOldBrush&)
                Call DeleteObject(hPatBrush&)
                Call DeleteObject(hOldBrush&)
            End If
            If bDisabled Then
                Call PrintGlyph(hdc&, IIf(bDrawCheck, "a", "h"), GetSysColor(COLOR_BTNHIGHLIGHT), OffsetRect(rc, 1, 1), DT_CENTER Or DT_VCENTER Or DT_SINGLELINE)
                Call PrintGlyph(hdc&, IIf(bDrawCheck, "a", "h"), GetSysColor(COLOR_GRAYTEXT), rc, DT_CENTER Or DT_VCENTER Or DT_SINGLELINE)
            Else
                Call PrintGlyph(hdc&, IIf(bDrawCheck, "a", "h"), WndObj.ForeColor, rc, DT_CENTER Or DT_VCENTER Or DT_SINGLELINE)
            End If
            Call DrawEdge(hdc&, rc, BDR_SUNKENOUTER, BF_RECT)
        End If
    Else
        'Bitmap argument is valid
    End If
    Draw3DMark = True
    Set WndObj = Nothing
    Exit Function
hError:
    Debug.Print Err.Number; Err.Description; " ( Draw3DMark )"
    Err.Clear
End Function
Private Function OnDrawMainMenu(hwnd As Long, lParam As Long, MousePosition As Long) As Long
    On Error GoTo NOPOP
    Dim hdc                     As Long
    Dim WndObj                  As clsWndCoolMenu: Set WndObj = WndCol(CStr(hwnd&))
    Dim hMenu                   As Long: hMenu& = GetMenu(hwnd&)
    Dim dwPapi                  As Double
    Dim Papi                    As POINTAPI
    Dim PopedIndex              As Long
    Dim info                    As MENUITEMINFO
    Dim MenuRect                As RECT
    Dim OldPopedRect            As RECT
    If WndObj.MainPopedIndex = -2 Then
        Set WndObj = Nothing
        Exit Function
    End If
    If MousePosition <> 5 And MousePosition > 0 Then GoTo NOPOP
    If MousePosition = 5 Then
        Set WndObj = Nothing
        Exit Function
    End If
    Papi.X = LoWord(lParam&)
    Papi.Y = HiWord(lParam&)
    Call CopyMemory(dwPapi, Papi, LenB(Papi))
    Dim MenuHitIndex As Long
    MenuHitIndex& = MenuItemFromPoint(hwnd&, hMenu&, dwPapi)
    If MenuHitIndex& = -1 Then GoTo NOPOP
    PopedIndex& = WndObj.MainPopedIndex
    If MenuHitIndex& = PopedIndex& Then
        Set WndObj = Nothing
        Exit Function
    End If
    info.cbSize = LenB(info)
    info.fMask = MIIM_TYPE
    Call GetMenuItemInfo(hMenu&, MenuHitIndex&, MF_BYPOSITION, info)
    If info.fType And (Not MFT_OWNERDRAW) Then GoTo NOPOP
    If PopedIndex& <> -1 Then GoSub DRAWFLAT
    Call GetMenuItemRect(hwnd&, hMenu&, MenuHitIndex&, MenuRect)
    WndObj.MainPopedIndex = MenuHitIndex&
    hdc& = GetDC(0&)
    Call DrawEdge(hdc&, MenuRect, BDR_RAISEDINNER, BF_RECT)
    Call ReleaseDC(hwnd&, hdc&)
    OnDrawMainMenu = True
    Set WndObj = Nothing
    Exit Function
NOPOP:
    If WndObj.MainPopedIndex > -1 Then
        GoSub DRAWFLAT
        WndObj.MainPopedIndex = -1
    End If
    Set WndObj = Nothing
    Exit Function
DRAWFLAT:
    Call GetMenuItemRect(hwnd&, hMenu&, CLng(WndObj.MainPopedIndex), OldPopedRect)
    hdc& = GetDC(0&)
    Call DrawEdge(hdc&, OldPopedRect, BDR_RAISEDINNER, BF_RECT Or BF_FLAT)
    Call ReleaseDC(hwnd&, hdc&)
    Return
End Function
Private Sub PrintGlyph(hdc As Long, Glyph As String, Color As Long, rt As RECT, ByVal wFormat As Long)
    Dim tLF                     As LogFont
    Dim hOldFont                As Long
    If m_MarlettFont& = 0& Then
        tLF.lfFaceName = "Marlett" + Chr(0)
        tLF.lfCharSet = SYMBOL_CHARSET
        tLF.lfHeight = 13
        m_MarlettFont& = CreateFontIndirect(tLF)
    End If
    Call SetBkMode(hdc&, TRANSPARENT)
    hOldFont& = SelectObject(hdc&, m_MarlettFont&)
    Call SetTextColor(hdc&, Color&)
    Call DrawText(hdc&, Glyph, 1, rt, wFormat&)
End Sub
Public Function WindowProc(ByVal hwnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    On Error GoTo ErrorHandle
    Dim Result                  As Long
    Select Case Msg&
        Case WM_SETCURSOR
            Call OnDrawMainMenu(hwnd&, 0&, LoWord(lParam&))
        Case WM_NCHITTEST
            Call OnDrawMainMenu(hwnd&, lParam&, -1)
        Case WM_NCMOUSEMOVE
        Case WM_MEASUREITEM
            If OnMeasureItem(hwnd&, lParam&) Then
                WindowProc = True
                Exit Function
            End If
        Case WM_DRAWITEM
            If OnDrawItem(hwnd&, lParam&) Then
                WindowProc = True
                Exit Function
            End If
        Case WM_INITMENUPOPUP
            m_SideBitmapWidth = 0
            Call CallWindowProc(WndCol(CStr(hwnd&)).PrevProc, ByVal hwnd&, ByVal Msg&, ByVal wParam&, ByVal lParam&)
            Call OnInitMenuPopup(hwnd&, wParam&, LoWord(lParam&), CBool(HiWord(lParam&)))
            WindowProc = 0&
            Exit Function
        Case WM_MENUCHAR
            Result = OnMenuChar(LoWord(wParam&), HiWord(wParam&), lParam&)
            If Result <> 0 Then
                WindowProc = Result
                Exit Function
            End If
        Case WM_MENUSELECT
            Call OnMenuSelect(hwnd&, LoWord(wParam&), HiWord(wParam&), lParam&)
        Case WM_WINDOWPOSCHANGED
            Dim OldH            As Long: OldH& = WndCol(CStr(hwnd&)).SCMainMenu
            If (OldH& <> 0&) And (OldH& <> -1&) And (GetMenu(hwnd&) <> OldH&) Then
                WndCol(CStr(hwnd&)).SCMainMenu = True
                Call ConvertMenu(hwnd&, GetMenu(hwnd&), 0&, False, True, True)
            End If
    End Select
Continue:
    WindowProc& = CallWindowProc(WndCol(CStr(hwnd&)).PrevProc, hwnd&, Msg&, wParam&, lParam&)
    Exit Function
ErrorHandle:
    'Debug.Print Err.Number; Err.Description; " WindowProc"
    Err.Clear
End Function
Private Function HiWord(LongIn As Long) As Integer
    Call CopyMemory(HiWord, ByVal VarPtr(LongIn) + 2, 2)
End Function
Private Function LoWord(LongIn As Long) As Integer
    Call CopyMemory(LoWord, LongIn, 2)
End Function
Private Function HiByte(WordIn As Integer) As Byte
    Call CopyMemory(HiByte, ByVal VarPtr(WordIn) + 1, 2)
End Function
Private Function LoByte(WordIn As Integer) As Byte
    Call CopyMemory(LoByte, WordIn, 2)
End Function
Private Function MakeLong(ByVal LoWord As Integer, ByVal HiWord As Integer) As Long
    MakeLong = CLng(LoWord)
    Call CopyMemory(ByVal VarPtr(MakeLong) + 2, HiWord, 2)
End Function
Private Function MakeWord(ByVal LoByte As Byte, ByVal HiByte As Byte) As Integer
    MakeWord = CInt(LoByte)
    Call CopyMemory(ByVal VarPtr(MakeWord) + 1, HiByte, 1)
End Function
Public Function ColorEmbossed(hwnd As Long, Optional Value As Variant) As Boolean
    On Error Resume Next
    If IsMissing(Value) Then
        ColorEmbossed = WndCol(CStr(hwnd&)).ColorEmbossed
    Else
        WndCol(CStr(hwnd&)).ColorEmbossed = Value
        ColorEmbossed = Value
    End If
End Function
Public Function ComplexChecks(hwnd As Long, Optional Value As Variant) As Boolean
    On Error Resume Next
    If IsMissing(Value) Then
        ComplexChecks = WndCol(CStr(hwnd&)).ComplexChecks
    Else
        WndCol(CStr(hwnd&)).ComplexChecks = Value
        ComplexChecks = Value
    End If
  
End Function
Public Function SelectColor(hwnd As Long, Optional Value As Variant) As Long
    On Error Resume Next
    If IsMissing(Value) Then
        SelectColor = WndCol(CStr(hwnd&)).SelectColor
    Else
        WndCol(CStr(hwnd&)).SelectColor = Value
        SelectColor = Value
    End If
End Function
Public Function RightToLeft(hwnd As Long, Optional Value As Variant) As Boolean
    On Error Resume Next
    If IsMissing(Value) Then
        RightToLeft = WndCol(CStr(hwnd&)).RightToLeft
    Else
        WndCol(CStr(hwnd&)).RightToLeft = Value
        RightToLeft = Value
    End If
End Function
Public Function FullSelect(hwnd As Long, Optional Value As Variant) As Boolean
    On Error Resume Next
    If IsMissing(Value) Then
        FullSelect = WndCol(CStr(hwnd&)).FullSelect
    Else
        WndCol(CStr(hwnd&)).FullSelect = Value
        FullSelect = Value
    End If
End Function
Public Function FontSize(hwnd As Long, Optional Value As Variant) As Long
    On Error Resume Next
    If IsMissing(Value) Then
        FontSize = WndCol(CStr(hwnd&)).FontSize
    Else
        WndCol(CStr(hwnd&)).FontSize = Value
        Call DrawMenuBar(hwnd&)
        FontSize = Value
    End If
End Function
Public Function ForeColor(hwnd As Long, Optional Value As Variant) As Long
    On Error Resume Next
    If IsMissing(Value) Then
        ForeColor = WndCol(CStr(hwnd&)).ForeColor
    Else
        WndCol(CStr(hwnd&)).ForeColor = Value
        Call DrawMenuBar(hwnd&)
        ForeColor = Value
    End If
End Function
Public Function FontName(hwnd As Long, Optional Value As Variant) As String
    On Error Resume Next
    If IsMissing(Value) Then
        FontName = WndCol(CStr(hwnd&)).FontName
    Else
        WndCol(CStr(hwnd&)).FontName = Value
        Call DrawMenuBar(hwnd&)
        FontName = Value
    End If
End Function
Public Sub MDIChildMenu(hwnd As Long)
    On Error Resume Next
    Dim ParentWnd               As Long: ParentWnd& = GetParent(GetParent(hwnd&))
    Dim WndObj                  As clsWndCoolMenu: Set WndObj = WndCol(CStr(ParentWnd&))
    If Not (WndObj Is Nothing) Then If WndObj.SCMainMenu Then WndObj.SCMainMenu = GetMenu(ParentWnd&)
    Set WndObj = Nothing
End Sub
Public Function MakeRop4(fore As RasterOpConstants, back As RasterOpConstants) As Long
    MakeRop4 = MakeLong(0, MakeWord(0, LoByte(LoWord(back)))) Or fore
End Function


