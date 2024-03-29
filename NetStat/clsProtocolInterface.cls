VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsProtocolInterface"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'This Class Module thanks to Coding Genius, http://www.pscode.com/vb/scripts/ShowCode.asp?txtCodeId=46567&lngWId=1
Private Type tPluginProtocols
    sID                             As Integer
    tProtocol                       As clsProtocol
End Type
Private PluginProtocols()           As tPluginProtocols
'Private PluginProtocols             As colProtocols
Private m_lngSocket                 As Long
Private Subclasser                  As cSuperClass
Private cIPHeader                   As clsIPHeader
Implements ISubclassingSink
Private Function P_UBound(aArr() As tPluginProtocols) As Long
    On Error GoTo ErrClear
    P_UBound = UBound(aArr)
    Exit Function
ErrClear:
    P_UBound = -1
    Err.Clear
End Function
Public Sub AddinProtocol(Plugin As Object, ProtocolName As String, UniqueID As Long)
    Dim NewProtocol As New clsProtocol
    With NewProtocol
        .ProtocolName = ProtocolName
        .ProtocolID = UniqueID
        Set .ProtocolPlugin = Plugin
        Set .ProtocolPlugin.ProtocolInterface = Me
    End With
    ReDim Preserve PluginProtocols(P_UBound(PluginProtocols) + 1)
    Set PluginProtocols(P_UBound(PluginProtocols)).tProtocol = NewProtocol
    PluginProtocols(P_UBound(PluginProtocols)).sID = CInt(UniqueID)
    Set NewProtocol = Nothing
End Sub
Public Function SendData(DestAddress As String, DestPort As Long, Packet() As Byte) As Long
    Dim SockAddr As sockaddr_in
    SockAddr = saZero
    SockAddr.sin_family = AF_INET
    SockAddr.sin_port = htons(DestPort)
    If SockAddr.sin_port = INVALID_SOCKET Then
        SendData = 0
        Exit Function
    End If
    SockAddr.sin_addr = GetHostByNameAlias(DestAddress)
    If SockAddr.sin_addr = INADDR_NONE Then
        SendData = 0
        Exit Function
    End If
    SendData = sendto(m_lngSocket, ByVal VarPtr(Packet(0)), UBound(Packet) + 1, 0, SockAddr, LenB(SockAddr))
End Function
Private Sub ISubclassingSink_After(lReturn As Long, ByVal hwnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long)
    'On Error GoTo ErrClear
    Dim lSocket                     As Long
    Dim BytesRecieved               As Long
    Dim ReadBuffer()                As Byte
    Dim IPH                         As IPHeader
    Dim X                           As Integer
    If uMsg = WinsockMessage Then
        lSocket = wParam
        Select Case lParam
            Case FD_READ
                ReDim ReadBuffer(1499) As Byte
                BytesRecieved = recv(lSocket, ByVal VarPtr(ReadBuffer(0)), 1500, 0)
                If BytesRecieved <= LenB(IPH) Then Exit Sub
                ReDim Preserve ReadBuffer(BytesRecieved - 1) As Byte
                CopyMemory IPH, ByVal VarPtr(ReadBuffer(0)), LenB(IPH)
                With cIPHeader
                    .Checksum = IntegerToUnsigned(ntohs(IPH.ip_checksum))
                    .DestAddress = IPH.ip_destaddr
                    .HeaderLength = LoNibble(IPH.ip_verlen) * 4
                    .ID = IntegerToUnsigned(IPH.ip_id)
                    .Offset = IntegerToUnsigned(ntohs(IPH.ip_offset))
                    .PacketLength = ntohs(IPH.ip_totallength)
                    .Protocol = IPH.ip_protocol
                    .SourceAddress = IPH.ip_srcaddr
                    .TimeToLive = IPH.ip_ttl
                    .Version = HiNibble(IPH.ip_verlen)
                End With
                'Had to change this.
                'I was getting a "Object Does not suporrt this....." error.
                X = P_UBound(PluginProtocols)
                For X = 0 To X
                    If PluginProtocols(X).sID = IPH.ip_protocol Then
                        PluginProtocols(X).tProtocol.ProtocolPlugin.PacketArrived cIPHeader, ReadBuffer(), BytesRecieved
                        Exit For
                    End If
                Next
                'For Each DummyProto In PluginProtocols
                    'If DummyProto.ProtocolID = IPH.ip_protocol Then
                        'DummyProto.ProtocolPlugin.PacketArrived cIPHeader, ReadBuffer(), BytesRecieved
                        'Exit For
                    'End If
                'Next
                lReturn = 1
            Case FD_WRITE
                lReturn = 1
        End Select
    End If
    Exit Sub
ErrClear:
    lReturn = 1
End Sub
Public Function CreateRawSocket(strAddress As String, lngPort As Long, hwnd As Long) As Long
    Dim RcvTimeOut                  As Long
    Dim SockAddr                    As sockaddr_in
    Dim lngInBuffer                 As Long
    Dim lngBytesReturned            As Long
    Dim lngOutBuffer                As Long
    SockAddr = saZero
    SockAddr.sin_family = AF_INET
    SockAddr.sin_port = htons(lngPort)
    If SockAddr.sin_port = INVALID_SOCKET Then
        CreateRawSocket = INVALID_SOCKET
        MsgBox "The port given is invalid."
        Exit Function
    End If
    SockAddr.sin_addr = GetHostByNameAlias(strAddress)
    If SockAddr.sin_addr = INADDR_NONE Then
        CreateRawSocket = INVALID_SOCKET
        MsgBox "The address given is invalid. Pass an existing IP of the form ###.###.###.### or a valid host name"
        Exit Function
    End If
    m_lngSocket = socket(AF_INET, SOCK_RAW, IPPROTO_IP)
    If m_lngSocket = INVALID_SOCKET Then
        CreateRawSocket = INVALID_SOCKET
        MsgBox "Could not create socket. 'socket(AF_INET, SOCK_RAW, IPPROTO_IP)' failed"
        Exit Function
    End If
    RcvTimeOut = 5000
    If setsockopt(m_lngSocket, SOL_SOCKET, SO_RCVTIMEO, RcvTimeOut, 4) <> 0 Then
        CreateRawSocket = INVALID_SOCKET
        MsgBox "Failed to set timeout. 'setsockopt(m_lngSocket, SOL_SOCKET, SO_RCVTIMEO, RcvTimeOut, 4)' failed"
        CloseSocket m_lngSocket
        Exit Function
    End If
    If Bind(m_lngSocket, SockAddr, LenB(SockAddr)) <> 0 Then
        CreateRawSocket = INVALID_SOCKET
        
        Debug.Print WSAGetLastError
        
        MsgBox "Failed to bind socket. 'bind(m_lngSocket, sockin, LenB(sockin))' failed"
        CloseSocket m_lngSocket
        Exit Function
    End If
    lngInBuffer = 1
    If WSAIoctl(m_lngSocket, SIO_RCVALL, lngInBuffer, Len(lngInBuffer), lngOutBuffer, Len(lngOutBuffer), lngBytesReturned, ByVal 0, ByVal 0) <> 0 Then
        CreateRawSocket = INVALID_SOCKET
        MsgBox "WSAIoctl failed."
        CloseSocket m_lngSocket
        Exit Function
    End If
    If WSAAsyncSelect(m_lngSocket, hwnd, WinsockMessage, ByVal FD_READ Or FD_WRITE) <> 0 Then
        CreateRawSocket = INVALID_SOCKET
        MsgBox "WSAAsyncSelect failed."
        CloseSocket m_lngSocket
        Exit Function
    End If
    Set Subclasser = New cSuperClass
    Subclasser.AddAfterMsgs (WinsockMessage)
    Subclasser.Subclass hwnd, Me
    CreateRawSocket = m_lngSocket
End Function
Public Function CloseRawSocket()
    If Not Subclasser Is Nothing Then Subclasser.Unsubclass
    Set Subclasser = Nothing
    CloseSocket m_lngSocket
End Function
Private Sub Class_Initialize()
    Dim StartupInfo                 As WSAData
    WSAStartup &H202, StartupInfo
    WinsockMessage = RegisterWindowMessage("RawSocketInterface")
    Set cIPHeader = New clsIPHeader
End Sub
Private Sub Class_Terminate()
    Erase PluginProtocols
    Set cIPHeader = Nothing
    If Not Subclasser Is Nothing Then
        Subclasser.Unsubclass
        Set Subclasser = Nothing
    End If
    CloseSocket m_lngSocket
    WSACleanup
End Sub
Private Sub ISubclassingSink_Before(bHandled As Boolean, lReturn As Long, hwnd As Long, uMsg As Long, wParam As Long, lParam As Long)
    'Unused
End Sub

