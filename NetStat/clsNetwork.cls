VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsNetwork"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
' Code built on MSKB Atricles
' Variouse sources on the internet
' And Basic Sockets/Winsock knowledge
' MSKB articles listed below:
' ID: Q160215
' ID: Q237688
' ID: Q194938
' ID: Q175472
' If Events are not used in this class then
' remove the following line
' and remove the withevent decl. from all other objects
Public Event DummyEvent(sDummy As Long)
Private Const NETWORK_ALIVE_LAN     As Long = &H1
Private Const NETWORK_ALIVE_WAN     As Long = &H2
Private Const NETWORK_ALIVE_AOL     As Long = &H4
Private Const lSocketErr            As Long = -1
Private Const AFInet                As Long = 2
Private Type HOSTENT
    lName                           As Long
    lAlias                          As Long
    iAddrType                       As Integer
    iLen                            As Integer
    lAddList                        As Long
End Type
Private Type QOCINFO
    dwSize                          As Long
    dwFlags                         As Long
    dwInSpeed                       As Long
    dwOutSpeed                      As Long
End Type
Private Declare Function IsDestinationReachable Lib "sensapi.dll" Alias "IsDestinationReachableA" (ByVal lpszDestination As String, lpQOCInfo As QOCINFO) As Long
Private Declare Function inet_addr Lib "wsock32" (ByVal s As String) As Long
Private Declare Function gethostname Lib "wsock32" (ByVal szHost As String, ByVal dwHostLen As Long) As Long
Private Declare Function gethostbyaddr Lib "wsock32" (haddr As Long, ByVal hnlen As Long, ByVal addrtype As Long) As Long
Private Declare Function gethostbyname Lib "wsock32" (ByVal hostname As String) As Long
Private Declare Function WSAGetLastError Lib "wsock32" () As Long
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (xDest As Any, xSource As Any, ByVal nbytes As Long)
Private Declare Function lstrlen Lib "kernel32" Alias "lstrlenA" (lpString As Any) As Long
Public Function GHostName(ByVal sIPAdd As String) As String
    Dim lHosent                     As Long
    Dim lAdd                        As Long
    Dim lBytes                      As Long
    Dim udtWSAData                  As WSAData
    If WSAStartup(257, udtWSAData) = 0 Then
        ' Convert address to long
        lAdd = inet_addr(sIPAdd)
        If lAdd <> lSocketErr Then
            ' Get pointer to the HOSTENT structure
            lHosent = gethostbyaddr(lAdd, 4, AFInet)
            ' If the pointer was returned
            If lHosent <> 0 Then
                CopyMemory lHosent, ByVal lHosent, 4
                lBytes = lstrlen(ByVal lHosent)
                If lBytes > 0 Then
                    sIPAdd = Space$(lBytes)
                    CopyMemory ByVal sIPAdd, ByVal lHosent, lBytes
                    GHostName = sIPAdd
                End If
            Else
                ' Host name not found
                GHostName = "Not Found"
            End If
            ' Close the open socket
            WSACleanup
        Else
            ' The IP Address passed seems to be invlaid
            MsgBox "IP Address Passed is Invalid!", vbOKOnly + vbExclamation, "GHostName Error"
            GHostName = "Not Found"
        End If
    Else
        ' Problem opening windows sockets
        ' user can try restarting windows
        MsgBox "Unable to initialize the Windows Sockets", vbOKOnly + vbInformation, "GHostName Error"
        GHostName = "Not Found"
    End If
End Function
Public Function GetIPAdd(ByVal sHostName As String) As String
    Dim lHosent                     As Long
    Dim lAdd                        As Long
    Dim lIPAdd                      As Long
    Dim sIPAdd                      As String
    Dim udtWSAData                  As WSAData
    If WSAStartup(257, udtWSAData) = 0 Then
        sIPAdd = Space(4)
        lHosent = gethostbyname(sHostName & vbNullChar)
        If lHosent <> 0 Then
            lAdd = lHosent + 12
            ' Get the IP address
            CopyMemory lAdd, ByVal lAdd, 4
            CopyMemory lIPAdd, ByVal lAdd, 4
            CopyMemory ByVal sIPAdd, ByVal lIPAdd, 4
            GetIPAdd = IPAdd2Text(sIPAdd)
        Else
            ' The IP Address was not found
            GetIPAdd = "IP Address Not Found!"
        End If
    End If
End Function
Private Function IPAdd2Text(ByVal sIPAdd As String) As String
    IPAdd2Text = CStr(Asc(sIPAdd)) & "." & CStr(Asc(Mid$(sIPAdd, 2, 1))) & "." & CStr(Asc(Mid$(sIPAdd, 3, 1))) & "." & CStr(Asc(Mid$(sIPAdd, 4, 1)))
End Function
Public Function CheckDest(sAddress As String) As Long
    Dim objQOC As QOCINFO
    ' Check to see if the address is reachable
    objQOC.dwSize = Len(objQOC)
    CheckDest = IsDestinationReachable(sAddress, objQOC)
End Function
Public Function GetLocalIPAddress() As String
    Dim sHostName                   As String * 256
    Dim lpHost                      As Long
    Dim HOST                        As HOSTENT
    Dim dwIPAddr                    As Long
    Dim tmpIPAddr()                 As Byte
    Dim i                           As Integer
    Dim sIPAddr                     As String
    Dim udtWSAData                  As WSAData
    If Not (WSAStartup(257, udtWSAData) = 0) Then
        GetLocalIPAddress = ""
        Exit Function
    End If
    If gethostname(sHostName, 256) = lSocketErr Then
        GetLocalIPAddress = ""
        MsgBox "Windows Sockets error " & Str$(WSAGetLastError()) & _
                " has occurred. Unable to successfully get Host Name."
        WSACleanup
        Exit Function
    End If
    sHostName = Trim$(sHostName)
    lpHost = gethostbyname(sHostName)
    If lpHost = 0 Then
        GetLocalIPAddress = ""
        MsgBox "Windows Sockets are not responding. " & _
            "Unable to successfully get Host Name."
        WSACleanup
        Exit Function
    End If
    CopyMemory HOST, lpHost, Len(HOST)
    CopyMemory dwIPAddr, HOST.lAddList, 4
    ReDim tmpIPAddr(1 To HOST.iLen)
    CopyMemory tmpIPAddr(1), dwIPAddr, HOST.iLen
    For i = 1 To HOST.iLen
        sIPAddr = sIPAddr & tmpIPAddr(i) & "."
    Next
    GetLocalIPAddress = Mid$(sIPAddr, 1, Len(sIPAddr) - 1)
    WSACleanup
End Function
Public Function GetLocalIPHostName() As String
    Dim sHostName                   As String * 256
    Dim udtWSAData                  As WSAData
    If WSAStartup(257, udtWSAData) = 1 Then
        GetLocalIPHostName = ""
        Exit Function
    End If
    If gethostname(sHostName, 256) = lSocketErr Then
        GetLocalIPHostName = ""
        MsgBox "There was an error retrieving the Local Host Name for your computer!", vbOKOnly + vbExclamation, "Sockets Error"
        WSACleanup
        Exit Function
    End If
    GetLocalIPHostName = Left$(sHostName, InStr(sHostName, Chr(0)) - 1)
    WSACleanup
End Function


