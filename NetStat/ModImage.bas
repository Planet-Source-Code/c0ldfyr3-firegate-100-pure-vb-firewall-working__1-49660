Attribute VB_Name = "ModImage"
Option Explicit
Private Declare Function CopyImage Lib "user32" (ByVal Handle As Long, ByVal un1 As Long, ByVal n1 As Long, ByVal n2 As Long, ByVal un2 As Long) As Long
Private Declare Function OleCreatePictureIndirect Lib "olepro32" (ByVal Pict As Long, riid As GUID, fOwn As Boolean, ppvObj As Object) As Long
Private Const IMAGE_BITMAP = 0&
Private Const IMAGE_ICON = 1
Private Const IMAGE_CURSOR = 2
Private Const IMAGE_ENHMETAFILE = 3
Private Const LR_COPYFROMRESOURCE = &H4000
Private Type PicBmp
    Size                            As Long
    tType                           As Long
    hBmp                            As Long
    hPal                            As Long
    Reserved                        As Long
End Type
Private Type PICTDESCBMP
    SizeofStruct                As Long
    PicType                     As Long
    hBitmap                     As Long
    hPalette                    As Long
End Type
Private Type PICTDESCICON
    SizeofStruct                As Long
    PicType                     As Long
    hIcon                       As Long
    lReserved                   As Long
End Type
Private Type GUID
    Data1                       As Long
    Data2                       As Integer
    Data3                       As Integer
    Data4(0& To 7)              As Byte
End Type
Private Function GetOlePicture(ByVal hImage As Long, ByVal ImageType As Long, Optional ByVal Own As Boolean) As StdPicture
    Dim varPicData              As PICTDESCBMP
    Dim varIconData             As PICTDESCICON
    Dim varIPicture             As StdPicture
    Dim varOLEID                As GUID
    Dim varhBmp                 As Long
    With varOLEID
        .Data1 = &H7BF80981
        .Data2 = &HBF32
        .Data3 = &H101A
        .Data4(0&) = &H8B
        .Data4(1) = &HBB
        .Data4(3) = &HAA
        .Data4(5) = &H30
        .Data4(6) = &HC
        .Data4(7) = &HAB
    End With
    If (ImageType = IMAGE_BITMAP) Then
        varPicData.SizeofStruct = Len(varPicData)
        varPicData.hBitmap = hImage
        varPicData.PicType = 1
        OleCreatePictureIndirect VarPtr(varPicData), varOLEID, Own, varIPicture
    ElseIf (ImageType = IMAGE_ICON) Or (ImageType = IMAGE_CURSOR) Then
        varIconData.SizeofStruct = Len(varIconData)
        varIconData.hIcon = hImage
        varIconData.PicType = 3
        OleCreatePictureIndirect ByVal VarPtr(varIconData), varOLEID, True, varIPicture
    End If
    On Error Resume Next
    If Not varIPicture Is Nothing Then
        Set GetOlePicture = varIPicture
    End If
End Function
Public Function ConvertTo16(hImage As Long) As StdPicture
    Dim Lng                     As Long
    Lng = CopyImage(hImage, IMAGE_ICON, 16, 16, LR_COPYFROMRESOURCE)
    Set ConvertTo16 = GetOlePicture(Lng, IMAGE_ICON, True)
End Function
