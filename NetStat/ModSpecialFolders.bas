Attribute VB_Name = "ModSpecialFolders"
Option Explicit
'Function to get any Special Folder contained on the windows OS.
Private Declare Function GetWindowsDirectory Lib "kernel32.dll" Alias "GetWindowsDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long
Private Declare Function GetSystemDirectory Lib "kernel32.dll" Alias "GetSystemDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long
Private Declare Function GetTempPath Lib "kernel32.dll" Alias "GetTempPathA" (ByVal nBufferLength As Long, ByVal lpBuffer As String) As Long
Private Declare Function SHGetSpecialFolderLocation Lib "shell32.dll" (ByVal hWndOwner As Long, ByVal nFolder As Long, pidl As ITEMIDLIST) As Long
Private Declare Function SHGetPathFromIDList Lib "shell32.dll" Alias "SHGetPathFromIDListA" (ByVal pidl As Long, ByVal pszPath As String) As Long
Private Type SHITEMID
    cb                              As Long
    abID                            As Byte
End Type
Private Type ITEMIDLIST
    mkid                            As SHITEMID
End Type
Public Enum Folders 'List of possible folders we may need.
    Windows = vbNull
    WinSystem = -1
    Desktop = &H0
    Internet = &H1
    Programs = &H2
    ControlsFolder = &H3
    Printers = &H4
    Documents = &H5
    Favorites = &H6
    StartUp = &H7
    Recent = &H8
    SendToFolder = &H9
    RecycleBin = &HA
    StartMenu = &HB
    DesktopDirectory = &H10
    Drives = &H11
    Network = &H12
    Nethood = &H13
    Fonts = &H14
    Templates = &H15
    Common_StartMenu = &H16
    Common_Programs = &H17
    Common_StartUp = &H18
    Common_DesktopDirectory = &H19
    ApplicationData = &H1A
    PrintHood = &H1B
    AltStartUp = &H1D
    Common_AltStartUp = &H1E
    Common_Favorites = &H1F
    InternetCache = &H20
    cookies = &H21
    History = &H22
    Temp = 99
End Enum
Public Function SpecialFolder(Optional TheFolder As Folders = vbNull) As String
    Dim ID                      As ITEMIDLIST
    Dim LngRet                  As Long
    Dim ThePath                 As String
    Dim TheLength               As Long
    ThePath = Space(255)
    Select Case TheFolder
        Case Windows
            'If they want the windows folder, get that on it's own.
            TheLength = GetWindowsDirectory(ThePath, 255)
            ThePath = Left(ThePath, TheLength)
        Case WinSystem
            'Get System folder.
            TheLength = GetSystemDirectory(ThePath, 255)
            ThePath = Left(ThePath, TheLength)
        Case Temp
            'Get Temp folder.
            TheLength = GetTempPath(255, ThePath)
            ThePath = Left(ThePath, TheLength)
        Case Else
            'Otherwise get the Special Folder requested.
            LngRet = SHGetSpecialFolderLocation(0, TheFolder, ID)
            If LngRet = 0 Then
                If SHGetPathFromIDList(ID.mkid.cb, ThePath) <> 0 Then
                    ThePath = Left(ThePath, InStr(ThePath, vbNullChar) - 1)
                End If
            End If
    End Select
    SpecialFolder = Trim(ThePath)
End Function
