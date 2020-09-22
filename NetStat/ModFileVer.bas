Attribute VB_Name = "ModFileVer"
Option Explicit
'Get File Version
Private Declare Function GetFileVersionInfoSize Lib "version.dll" Alias "GetFileVersionInfoSizeA" (ByVal lptstrFilename As String, lpdwHandle As Long) As Long
Private Declare Function GetFileVersionInfo Lib "version.dll" Alias "GetFileVersionInfoA" (ByVal lptstrFilename As String, ByVal dwHandle As Long, ByVal dwLen As Long, lpData As Any) As Long
Public Enum eFileInformation 'I like to use Enum's alot, they make the code easier to use =)
    [Company Name] = 0
    [File Description] = 1
    [Internal Name] = 2
    [Legal Copyright] = 3
    [Legal Trademarks] = 4
    [Original FileName] = 5
    [Product Name] = 6
    [Product Version] = 7
    [Comments] = 8
    [File Version] = 9
End Enum
Private Function sInfoType(tType As eFileInformation) As String
    'Return the string to search for.
    Select Case tType
        Case Is = 0
            sInfoType = "CompanyName"
        Case Is = 1
            sInfoType = "FileDescription"
        Case Is = 2
            sInfoType = "InternalName"
        Case Is = 3
            sInfoType = "LegalCopyright"
        Case Is = 4
            sInfoType = "LegalTrademarks"
        Case Is = 5
            sInfoType = "OriginalFileName"
        Case Is = 6
            sInfoType = "ProductName"
        Case Is = 7
            sInfoType = "ProductVersion"
        Case Is = 8
            sInfoType = "Comments"
        Case Is = 9
            sInfoType = "FileVersion"
    End Select
End Function
Public Function FileInfo(FileName As String, tType As eFileInformation) As String
    Dim lHwnd               As Long
    Dim lLen                As Long
    Dim bBuffer()           As Byte
    Dim sText               As String
    Dim iPos(1)             As Integer
    lLen = GetFileVersionInfoSize(FileName, lHwnd) 'Get the Version header.
    If lLen <= 1 Then Exit Function 'If there is no Version header.
    ReDim bBuffer(lLen)
    If GetFileVersionInfo(FileName, lHwnd, lLen, bBuffer(0)) = 1 Then 'Grab the header.
        sText = StrConv(StrConv(bBuffer(), vbUnicode), vbFromUnicode) 'Convert it to legible text.
        iPos(0) = InStr(1, sText, sInfoType(tType)) 'Search for the type of information you want.
        If iPos(0) = 0 Then Exit Function 'If it doesn't contain the type we want.
        iPos(0) = InStr(iPos(0), sText, Chr(0)) + 1
        iPos(1) = InStr(iPos(0), sText, Chr(0)) + 1
        If iPos(0) + 1 = iPos(1) Then
            iPos(0) = iPos(0) + 1
            iPos(1) = InStr(iPos(0), sText, Chr(0))
        End If
        FileInfo = Mid(sText, iPos(0), iPos(1) - iPos(0)) 'Capture the data.
    End If
End Function
