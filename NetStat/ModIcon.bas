Attribute VB_Name = "ModIcon"
Option Explicit
Private Declare Function DestroyIcon Lib "user32" (ByVal hIcon As Long) As Long
Private Declare Function ExtractIconEx Lib "shell32.dll" Alias "ExtractIconExA" (ByVal lpszFile As String, ByVal nIconIndex As Long, phiconLarge As Long, phiconSmall As Long, ByVal nIcons As Long) As Long
Private Declare Function ImageList_Draw Lib "comctl32.dll" (ByVal himl&, ByVal i&, ByVal hDCDest&, ByVal X&, ByVal Y&, ByVal Flags&) As Long
Private Declare Function OleCreatePictureIndirect Lib "olepro32.dll" (PicDesc As PicBmp, RefIID As GUID, ByVal fPictureOwnsHandle As Long, IPic As IPicture) As Long
Private Declare Function SHGetFileInfo Lib "shell32.dll" Alias "SHGetFileInfoA" (ByVal pszPath As String, ByVal dwFileAttributes As Long, psfi As SHFILEINFO, ByVal cbSizeFileInfo As Long, ByVal uFlags As Long) As Long
Public Const LARGE_ICON             As Integer = 32
Public Const SMALL_ICON             As Integer = 16
Private Const MAX_PATH              As Integer = 400
Private Const SHGFI_SYSICONINDEX    As Long = &H4000
Private Const SHGFI_SHELLICONSIZE   As Long = &H4
Private Const SHGFI_DISPLAYNAME     As Long = &H200
Private Const SHGFI_EXETYPE         As Long = &H2000
Private Const SHGFI_LARGEICON       As Long = &H0
Private Const SHGFI_SMALLICON       As Long = &H1
Private Const SHGFI_TYPENAME        As Long = &H400
Private Const ILD_TRANSPARENT       As Long = &H1
Private Const BASIC_SHGFI_FLAGS     As Long = SHGFI_TYPENAME Or SHGFI_SHELLICONSIZE Or SHGFI_SYSICONINDEX Or SHGFI_DISPLAYNAME Or SHGFI_EXETYPE
Private Type SHFILEINFO
    hIcon                           As Long
    iIcon                           As Long
    dwAttributes                    As Long
    szDisplayName                   As String * MAX_PATH
    szTypeName                      As String * 80
End Type
Private Type tImages
    PicLargeIcon                    As Picture
    PicSmallIcon                    As Picture
End Type
Private Type GUID
    Data1                           As Long
    Data2                           As Integer
    Data3                           As Integer
    Data4(7)                        As Byte
End Type
Private Type PicBmp
    Size                            As Long
    tType                           As Long
    hBmp                            As Long
    hPal                            As Long
    Reserved                        As Long
End Type
Public Type tPicIndex
    Pic16                           As Integer
    Pic32                           As Integer
End Type
Public Function FileIconToPicture(FileName As String, PicBox As PictureBox, Img As Image)
    Dim hLIcon                      As Long
    Dim ShInfo                      As SHFILEINFO
    hLIcon = SHGetFileInfo(FileName, 0&, ShInfo, Len(ShInfo), BASIC_SHGFI_FLAGS Or SHGFI_LARGEICON)
    If Not hLIcon = 0 Then
        'Locked onto an icon...
        With PicBox
            Set .Picture = LoadPicture("") 'Clear the existing pic.
            .AutoRedraw = True
            Call ImageList_Draw(hLIcon, ShInfo.iIcon, PicBox.hdc, 0, 0, ILD_TRANSPARENT)
            'Draw the image into the Pic Box
            .Refresh
        End With
        Img.Picture = PicBox.Image
    End If
End Function
Public Function GetIcon(FileName As String, Index As Long, sShell32 As String, Pic16 As PictureBox, Pic32 As PictureBox, Img32 As ImageList, Img16 As ImageList) As tPicIndex
    Dim hLIcon                      As Long
    Dim hSIcon                      As Long
    Dim imgObj                      As ListImage
    Dim r                           As Long
    Dim lIcons                      As tImages
    Dim ShInfo                      As SHFILEINFO
    'hLIcon = SHGetFileInfo(FileName, FILE_ATTRIBUTE_NORMAL, ShInfo, Len(ShInfo), SH_USEFILEATTRIBUTES Or BASIC_SH_FLAGS Or SH_LARGEICON)
    'hSIcon = SHGetFileInfo(FileName, FILE_ATTRIBUTE_NORMAL, ShInfo, Len(ShInfo), SH_USEFILEATTRIBUTES Or BASIC_SH_FLAGS Or SH_SMALLICON)
    hSIcon = SHGetFileInfo(FileName, 0&, ShInfo, Len(ShInfo), BASIC_SHGFI_FLAGS Or SHGFI_SMALLICON)
    hLIcon = SHGetFileInfo(FileName, 0&, ShInfo, Len(ShInfo), BASIC_SHGFI_FLAGS Or SHGFI_LARGEICON)
    'Get the file info
    GetIcon.Pic16 = Img16.ListImages.Count + 1
    GetIcon.Pic32 = Img32.ListImages.Count + 1
    If Not hLIcon = 0 Then
        'Locked onto an icon...
        With Pic32
            Set .Picture = LoadPicture("") 'Clear the existing pic.
            .AutoRedraw = True
            r = ImageList_Draw(hLIcon, ShInfo.iIcon, Pic32.hdc, 0, 0, ILD_TRANSPARENT)
            'Draw the image into the Pic Box
            .Refresh
        End With
        With Pic16
            Set .Picture = LoadPicture("") 'Clear the existing pic.
            .AutoRedraw = True
            r = ImageList_Draw(hSIcon, ShInfo.iIcon, Pic16.hdc, 0, 0, ILD_TRANSPARENT)
            'Draw the image into the Pic Box
            .Refresh
        End With
        'Add it to the image list
        Set imgObj = Img32.ListImages.Add(GetIcon.Pic32, , Pic32.Image)
        Set imgObj = Img16.ListImages.Add(GetIcon.Pic16, , Pic16.Image)
    Else
        lIcons = GetIconFromFile(sShell32, 0, True)
        With Pic32
            .AutoRedraw = True
            Set .Picture = lIcons.PicLargeIcon
            .Refresh
        End With
        With Pic16
            .AutoRedraw = True
            Set .Picture = lIcons.PicSmallIcon  'Get the default "No Icon" icon from shell32.dll in %WinDir%\System32
            .Refresh
        End With
        Set imgObj = Img32.ListImages.Add(GetIcon.Pic32, , Pic32.Image)
        Set imgObj = Img16.ListImages.Add(GetIcon.Pic16, , Pic16.Image)
    End If
End Function
Public Function GetIconFromFile(FileName As String, IconIndex As Long, UseLargeIcon As Boolean) As tImages
    Dim hLargeIcon                  As Long
    Dim hSmallIcon                  As Long
    Dim pic(1)                      As PicBmp
    Dim IPic(1)                     As IPicture
    Dim IID_IDispatch(1)            As GUID
    If ExtractIconEx(FileName, IconIndex, hLargeIcon, hSmallIcon, 1) > 0 Then
        With IID_IDispatch(0)
            .Data1 = &H20400
            .Data4(0) = &HC0
            .Data4(7) = &H46
        End With
        With IID_IDispatch(1)
            .Data1 = &H20400
            .Data4(0) = &HC0
            .Data4(7) = &H46
        End With
        With pic(0)
            .Size = Len(pic(0))
            .tType = vbPicTypeIcon
            .hBmp = hLargeIcon
        End With
        With pic(1)
            .Size = Len(pic(1))
            .tType = vbPicTypeIcon
            .hBmp = hSmallIcon
        End With
        Call OleCreatePictureIndirect(pic(0), IID_IDispatch(0), 1, IPic(0))
        Call OleCreatePictureIndirect(pic(1), IID_IDispatch(1), 1, IPic(1))
        Set GetIconFromFile.PicLargeIcon = IPic(0)
        Set GetIconFromFile.PicSmallIcon = IPic(1)
        DestroyIcon hSmallIcon
        DestroyIcon hLargeIcon
    End If
End Function
