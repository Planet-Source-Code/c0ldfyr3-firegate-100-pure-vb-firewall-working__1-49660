Attribute VB_Name = "ModSocks"
Option Explicit
'********************************************************************************
'MSocketSupport module
'Copyright 2002 by Oleg Gdalevich
'Visual Basic Internet Programming website (http://www.vbip.com)
'********************************************************************************
'This module contains API declarations and helper functions for the CSocket class
'********************************************************************************
'Version: 1.0.12     Modified: 17-OCT-2002
'********************************************************************************
'To get latest version of this code please visit the following web page:
'http://www.vbip.com/winsock-api/csocket-class/csocket-class-01.asp
'********************************************************************************
Public Const INADDR_NONE            As Long = &HFFFF
Public Const SOCKET_ERROR           As Long = -1
Public Const INVALID_SOCKET         As Long = -1
Public Const INADDR_ANY             As Long = &H0
Public Const FD_SETSIZE             As Long = 64
Public Const MAXGETHOSTSTRUCT       As Long = 1024
Public Const SD_RECEIVE             As Long = &H0
Public Const SD_SEND                As Long = &H1
Public Const SD_BOTH                As Long = &H2
Public Const MSG_OOB                As Long = &H1               '/* process out-of-band data */
Public Const MSG_PEEK               As Long = &H2               '/* peek at incoming message */
Public Const MSG_DONTROUTE          As Long = &H4               '/* send without using routing tables */
Public Const MSG_PARTIAL            As Long = &H8000            '/* partial send or recv for message xport */
Public Const FD_READ                As Long = &H1&
Public Const FD_WRITE               As Long = &H2&
Public Const FD_OOB                 As Long = &H4&
Public Const FD_ACCEPT              As Long = &H8&
Public Const FD_CONNECT             As Long = &H10&
Public Const FD_CLOSE               As Long = &H20&
Public Const SOL_SOCKET             As Long = 65535
Public Const SO_SNDBUF              As Long = &H1001            'Send buffer size
Public Const SO_RCVBUF              As Long = &H1002            'Receive buffer size
Public Const SO_SNDLOWAT            As Long = &H1003            'Send low-water mark
Public Const SO_RCVLOWAT            As Long = &H1004            'Receive low-water mark
Public Const SO_SNDTIMEO            As Long = &H1005            'Send timeout
Public Const SO_RCVTIMEO            As Long = &H1006            'Receive timeout
Public Const SO_ERROR               As Long = &H1007            'Get error status and clear
Public Const SO_TYPE                As Long = &H1008            'Get socket type
Public Const SO_DEBUG               As Long = &H1&              ' Turn on debugging info recording
Public Const SO_ACCEPTCONN          As Long = &H2&              ' Socket has had listen() - READ-ONLY.
Public Const SO_REUSEADDR           As Long = &H4&              ' Allow local address reuse.
Public Const SO_KEEPALIVE           As Long = &H8&              ' Keep connections alive.
Public Const SO_DONTROUTE           As Long = &H10&             ' Just use interface addresses.
Public Const SO_BROADCAST           As Long = &H20&             ' Permit sending of broadcast msgs.
Public Const SO_USELOOPBACK         As Long = &H40&             ' Bypass hardware when possible.
Public Const SO_LINGER              As Long = &H80&             ' Linger on close if data present.
Public Const SO_OOBINLINE           As Long = &H100&            ' Leave received OOB data in line.
Public Const SO_DONTLINGER          As Long = Not SO_LINGER
Public Const SO_EXCLUSIVEADDRUSE    As Long = Not SO_REUSEADDR  ' Disallow local address reuse.
Public Const WSADESCRIPTION_LEN     As Long = 257
Public Const WSASYS_STATUS_LEN      As Long = 129
Public Const SIO_GET_INTERFACE_LIST As Long = &H4004747F
Public Const TCP_NODELAY            As Long = &H1
Public Const TCP_BSDURGENT          As Long = &H7000
Public Type WSAData
    wVersion                        As Integer
    wHighVersion                    As Integer
    szDescription                   As String * WSADESCRIPTION_LEN
    szSystemStatus                  As String * WSASYS_STATUS_LEN
    iMaxSockets                     As Integer
    iMaxUdpDg                       As Integer
    lpVendorInfo                    As Long
End Type
Public Type sockaddr_in
    sin_family                      As Integer
    sin_port                        As Integer
    sin_addr                        As Long
    sin_zero(1 To 8)                As Byte
End Type
Public Type fd_set
  fd_count                          As Long '// how many are SET?
  fd_array(1 To FD_SETSIZE)         As Long '// an array of SOCKETs
End Type
Public Type IPHeader
    ip_verlen                       As Byte             'IP version number and header length in 32bit words (4 bits each)
    ip_tos                          As Byte             'Type Of Service ID (1 octet)
    ip_totallength                  As Integer          'Size of Datagram (header + data) in octets
    ip_id                           As Integer          'IP-ID (16 bits)
    ip_offset                       As Integer          'fragmentation flags (3bit) and fragmet offset (13 bits)
    ip_ttl                          As Byte             'datagram Time To Live (in network hops)
    ip_protocol                     As Byte             'Transport protocol type (byte)
    ip_checksum                     As Integer          'Header Checksum (16 bits)
    ip_srcaddr                      As Long             'Source IP Address (32 bits)
    ip_destaddr                     As Long             'Destination IP Address (32 bits)
End Type

Public Type sockaddr_gen
    AddressIn                       As sockaddr_in
    filler(0 To 7)                  As Byte
End Type
Public Type INTERFACE_INFO
    iiFlags                         As Long             'Interface flags
    iiAddress                       As sockaddr_gen     'Interface address
    iiBroadcastAddress              As sockaddr_gen     'Broadcast address
    iiNetmask                       As sockaddr_gen     'Network mask
End Type
Public Type INTERFACEINFO
    iInfo(0 To 7)                   As INTERFACE_INFO
End Type
Public Const WSABASEERR             As Long = 10000
Public Const WSAEINTR               As Long = (WSABASEERR + 4)
Public Const WSAEBADF               As Long = (WSABASEERR + 9)
Public Const WSAEACCES              As Long = (WSABASEERR + 13)
Public Const WSAEFAULT              As Long = (WSABASEERR + 14)
Public Const WSAEINVAL              As Long = (WSABASEERR + 22)
Public Const WSAEMFILE              As Long = (WSABASEERR + 24)
Public Const WSAEWOULDBLOCK         As Long = (WSABASEERR + 35)
Public Const WSAEINPROGRESS         As Long = (WSABASEERR + 36)
Public Const WSAEALREADY            As Long = (WSABASEERR + 37)
Public Const WSAENOTSOCK            As Long = (WSABASEERR + 38)
Public Const WSAEDESTADDRREQ        As Long = (WSABASEERR + 39)
Public Const WSAEMSGSIZE            As Long = (WSABASEERR + 40)
Public Const WSAEPROTOTYPE          As Long = (WSABASEERR + 41)
Public Const WSAENOPROTOOPT         As Long = (WSABASEERR + 42)
Public Const WSAEPROTONOSUPPORT     As Long = (WSABASEERR + 43)
Public Const WSAESOCKTNOSUPPORT     As Long = (WSABASEERR + 44)
Public Const WSAEOPNOTSUPP          As Long = (WSABASEERR + 45)
Public Const WSAEPFNOSUPPORT        As Long = (WSABASEERR + 46)
Public Const WSAEAFNOSUPPORT        As Long = (WSABASEERR + 47)
Public Const WSAEADDRINUSE          As Long = (WSABASEERR + 48)
Public Const WSAEADDRNOTAVAIL       As Long = (WSABASEERR + 49)
Public Const WSAENETDOWN            As Long = (WSABASEERR + 50)
Public Const WSAENETUNREACH         As Long = (WSABASEERR + 51)
Public Const WSAENETRESET           As Long = (WSABASEERR + 52)
Public Const WSAECONNABORTED        As Long = (WSABASEERR + 53)
Public Const WSAECONNRESET          As Long = (WSABASEERR + 54)
Public Const WSAENOBUFS             As Long = (WSABASEERR + 55)
Public Const WSAEISCONN             As Long = (WSABASEERR + 56)
Public Const WSAENOTCONN            As Long = (WSABASEERR + 57)
Public Const WSAESHUTDOWN           As Long = (WSABASEERR + 58)
Public Const WSAETOOMANYREFS        As Long = (WSABASEERR + 59)
Public Const WSAETIMEDOUT           As Long = (WSABASEERR + 60)
Public Const WSAECONNREFUSED        As Long = (WSABASEERR + 61)
Public Const WSAELOOP               As Long = (WSABASEERR + 62)
Public Const WSAENAMETOOLONG        As Long = (WSABASEERR + 63)
Public Const WSAEHOSTDOWN           As Long = (WSABASEERR + 64)
Public Const WSAEHOSTUNREACH        As Long = (WSABASEERR + 65)
Public Const WSAENOTEMPTY           As Long = (WSABASEERR + 66)
Public Const WSAEPROCLIM            As Long = (WSABASEERR + 67)
Public Const WSAEUSERS              As Long = (WSABASEERR + 68)
Public Const WSAEDQUOT              As Long = (WSABASEERR + 69)
Public Const WSAESTALE              As Long = (WSABASEERR + 70)
Public Const WSAEREMOTE             As Long = (WSABASEERR + 71)
Public Const WSASYSNOTREADY         As Long = (WSABASEERR + 91)
Public Const WSAVERNOTSUPPORTED     As Long = (WSABASEERR + 92)
Public Const WSANOTINITIALISED      As Long = (WSABASEERR + 93)
Public Const WSAEDISCON             As Long = (WSABASEERR + 101)
Public Const WSAENOMORE             As Long = (WSABASEERR + 102)
Public Const WSAECANCELLED          As Long = (WSABASEERR + 103)
Public Const WSAEINVALIDPROCTABLE   As Long = (WSABASEERR + 104)
Public Const WSAEINVALIDPROVIDER    As Long = (WSABASEERR + 105)
Public Const WSAEPROVIDERFAILEDINIT As Long = (WSABASEERR + 106)
Public Const WSASYSCALLFAILURE      As Long = (WSABASEERR + 107)
Public Const WSASERVICE_NOT_FOUND   As Long = (WSABASEERR + 108)
Public Const WSATYPE_NOT_FOUND      As Long = (WSABASEERR + 109)
Public Const WSA_E_NO_MORE          As Long = (WSABASEERR + 110)
Public Const WSA_E_CANCELLED        As Long = (WSABASEERR + 111)
Public Const WSAEREFUSED            As Long = (WSABASEERR + 112)
Public Const WSAHOST_NOT_FOUND      As Long = (WSABASEERR + 1001)
Public Const WSATRY_AGAIN           As Long = (WSABASEERR + 1002)
Public Const WSANO_RECOVERY         As Long = (WSABASEERR + 1003)
Public Const WSANO_DATA             As Long = (WSABASEERR + 1004)
Public Const SIO_RCVALL             As Long = &H98000001
Public Enum ServicePort
    IPPORT_ECHO = 7
    IPPORT_DISCARD = 9
    IPPORT_SYSTAT = 11
    IPPORT_DAYTIME = 13
    IPPORT_NETSTAT = 15
    IPPORT_FTP = 21
    IPPORT_TELNET = 23
    IPPORT_SMTP = 25
    IPPORT_TIMESERVER = 37
    IPPORT_NAMESERVER = 42
    IPPORT_WHOIS = 43
    IPPORT_MTP = 57
End Enum
Public Enum SocketType
    SOCK_STREAM = 1                 ' /* stream socket */
    SOCK_DGRAM = 2                  ' /* datagram socket */
    SOCK_RAW = 3                    ' /* raw-protocol interface */
    SOCK_RDM = 4                    ' /* reliably-delivered message */
    SOCK_SEQPACKET = 5              ' /* sequenced packet stream */
End Enum
Public Enum AddressFamily
    AF_UNSPEC = 0                   '/* unspecified */
    AF_UNIX = 1                     '/* local to host (pipes, portals) */
    AF_INET = 2                     '/* internetwork: UDP, TCP, etc. */
    AF_IMPLINK = 3                  '/* arpanet imp addresses */
    AF_PUP = 4                      '/* pup protocols: e.g. BSP */
    AF_CHAOS = 5                    '/* mit CHAOS protocols */
    AF_NS = 6                       '/* XEROX NS protocols */
    AF_IPX = AF_NS                  '/* IPX protocols: IPX, SPX, etc. */
    AF_ISO = 7                      '/* ISO protocols */
    AF_OSI = AF_ISO                 '/* OSI is ISO */
    AF_ECMA = 8                     '/* european computer manufacturers */
    AF_DATAKIT = 9                  '/* datakit protocols */
    AF_CCITT = 10                   '/* CCITT protocols, X.25 etc */
    AF_SNA = 11                     '/* IBM SNA */
    AF_DECnet = 12                  '/* DECnet */
    AF_DLI = 13                     '/* Direct data link interface */
    AF_LAT = 14                     '/* LAT */
    AF_HYLINK = 15                  '/* NSC Hyperchannel */
    AF_APPLETALK = 16               '/* AppleTalk */
    AF_NETBIOS = 17                 '/* NetBios-style addresses */
    AF_VOICEVIEW = 18               '/* VoiceView */
    AF_FIREFOX = 19                 '/* Protocols from Firefox */
    AF_UNKNOWN1 = 20                '/* Somebody is using this! */
    AF_BAN = 21                     '/* Banyan */
    AF_ATM = 22                     '/* Native ATM Services */
    AF_INET6 = 23                   '/* Internetwork Version 6 */
    AF_CLUSTER = 24                 '/* Microsoft Wolfpack */
    AF_12844 = 25                   '/* IEEE 1284.4 WG AF */
    AF_MAX = 26
End Enum
Public Const PF_UNSPEC              As Long = AF_UNSPEC
Public Const PF_UNIX                As Long = AF_UNIX
Public Const PF_INET                As Long = AF_INET
Public Const PF_IMPLINK             As Long = AF_IMPLINK
Public Const PF_PUP                 As Long = AF_PUP
Public Const PF_CHAOS               As Long = AF_CHAOS
Public Const PF_NS                  As Long = AF_NS
Public Const PF_IPX                 As Long = AF_IPX
Public Const PF_ISO                 As Long = AF_ISO
Public Const PF_OSI                 As Long = AF_OSI
Public Const PF_ECMA                As Long = AF_ECMA
Public Const PF_DATAKIT             As Long = AF_DATAKIT
Public Const PF_CCITT               As Long = AF_CCITT
Public Const PF_SNA                 As Long = AF_SNA
Public Const PF_DECnet              As Long = AF_DECnet
Public Const PF_DLI                 As Long = AF_DLI
Public Const PF_LAT                 As Long = AF_LAT
Public Const PF_HYLINK              As Long = AF_HYLINK
Public Const PF_APPLETALK           As Long = AF_APPLETALK
Public Const PF_VOICEVIEW           As Long = AF_VOICEVIEW
Public Const PF_FIREFOX             As Long = AF_FIREFOX
Public Const PF_UNKNOWN1            As Long = AF_UNKNOWN1
Public Const PF_BAN                 As Long = AF_BAN
Public Const PF_MAX                 As Long = AF_MAX
Public Enum SocketProtocol
    IPPROTO_IP = 0                  '/* dummy for IP */
    IPPROTO_ICMP = 1                '/* control message protocol */
    IPPROTO_IGMP = 2                '/* internet group management protocol */
    IPPROTO_GGP = 3                 '/* gateway^2 (deprecated) */
    IPPROTO_TCP = 6                 '/* tcp */
    IPPROTO_PUP = 12                '/* pup */
    IPPROTO_UDP = 17                '/* user datagram protocol */
    IPPROTO_IDP = 22                '/* xns idp */
    IPPROTO_ND = 77                 '/* UNOFFICIAL net disk proto */
    IPPROTO_RAW = 255               '/* raw IP packet */
    IPPROTO_MAX = 256
End Enum
Public Type HOSTENT
    hName                           As Long
    hAliases                        As Long
    hAddrType                       As Integer
    hLength                         As Integer
    hAddrList                       As Long
End Type
Public Declare Function gethostbyaddr Lib "ws2_32.dll" (addr As Long, ByVal addr_len As Long, ByVal addr_type As Long) As Long
Public Declare Function gethostbyname Lib "ws2_32.dll" (ByVal host_name As String) As Long
Public Declare Function gethostname Lib "ws2_32.dll" (ByVal host_name As String, ByVal namelen As Long) As Long
Public Declare Function getservbyname Lib "ws2_32.dll" (ByVal serv_name As String, ByVal proto As String) As Long
Public Declare Function getprotobynumber Lib "ws2_32.dll" (ByVal proto As Long) As Long
Public Declare Function getprotobyname Lib "ws2_32.dll" (ByVal proto_name As String) As Long
Public Declare Function getservbyport Lib "ws2_32.dll" (ByVal Port As Integer, ByVal proto As Long) As Long
Public Declare Function inet_addr Lib "ws2_32.dll" (ByVal cp As String) As Long
Public Declare Function inet_ntoa Lib "ws2_32.dll" (ByVal inn As Long) As Long
Public Declare Function htons Lib "ws2_32.dll" (ByVal hostshort As Integer) As Integer
Public Declare Function htonl Lib "ws2_32.dll" (ByVal hostlong As Long) As Long
Public Declare Function ntohl Lib "ws2_32.dll" (ByVal netlong As Long) As Long
Public Declare Function ntohs Lib "ws2_32.dll" (ByVal netshort As Integer) As Integer
Public Declare Function api_socket Lib "ws2_32.dll" Alias "socket" (ByVal af As Long, ByVal s_type As Long, ByVal Protocol As Long) As Long
Public Declare Function api_closesocket Lib "ws2_32.dll" Alias "closesocket" (ByVal s As Long) As Long
Public Declare Function api_connect Lib "ws2_32.dll" Alias "connect" (ByVal s As Long, ByRef Name As sockaddr_in, ByVal namelen As Long) As Long
Public Declare Function getsockname Lib "ws2_32.dll" (ByVal s As Long, ByRef Name As sockaddr_in, ByRef namelen As Long) As Long
Public Declare Function getpeername Lib "ws2_32.dll" (ByVal s As Long, ByRef Name As sockaddr_in, ByRef namelen As Long) As Long
Public Declare Function api_bind Lib "ws2_32.dll" Alias "bind" (ByVal s As Long, ByRef Name As sockaddr_in, ByRef namelen As Long) As Long
Public Declare Function api_select Lib "ws2_32.dll" Alias "select" (ByVal nfds As Long, ByRef readfds As Any, ByRef writefds As Any, ByRef exceptfds As Any, ByRef Timeout As Long) As Long
Public Declare Function recv Lib "ws2_32.dll" (ByVal s As Long, ByRef buf As Any, ByVal buflen As Long, ByVal Flags As Long) As Long
Public Declare Function send Lib "ws2_32.dll" (ByVal s As Long, ByRef buf As Any, ByVal buflen As Long, ByVal Flags As Long) As Long
Public Declare Function shutdown Lib "ws2_32.dll" (ByVal s As Long, ByVal how As Long) As Long
Public Declare Function api_listen Lib "ws2_32.dll" Alias "listen" (ByVal s As Long, ByVal backlog As Long) As Long
Public Declare Function api_accept Lib "ws2_32.dll" Alias "accept" (ByVal s As Long, ByRef addr As sockaddr_in, ByRef addrlen As Long) As Long
Public Declare Function setsockopt Lib "ws2_32.dll" (ByVal s As Long, ByVal level As Long, ByVal optname As Long, optval As Any, ByVal optlen As Long) As Long
Public Declare Function getsockopt Lib "ws2_32.dll" (ByVal s As Long, ByVal level As Long, ByVal optname As Long, optval As Any, optlen As Long) As Long
Public Declare Function sendto Lib "ws2_32.dll" (ByVal s As Long, ByRef buf As Any, ByVal buflen As Long, ByVal Flags As Long, ByRef toaddr As sockaddr_in, ByVal tolen As Long) As Long
Public Declare Function socket Lib "ws2_32.dll" (ByVal af As Long, ByVal s_type As Long, ByVal Protocol As Long) As Long
Public Declare Function Bind Lib "ws2_32.dll" Alias "bind" (ByVal s As Long, ByRef Name As sockaddr_in, ByRef namelen As Long) As Long
Public Declare Function CloseSocket Lib "ws2_32.dll" Alias "closesocket" (ByVal s As Long) As Long
Public Declare Function recvfrom Lib "ws2_32.dll" (ByVal s As Long, ByRef buf As Any, ByVal buflen As Long, ByVal Flags As Long, ByRef from As sockaddr_in, ByRef fromlen As Long) As Long
Public Declare Function WSAAsyncSelect Lib "ws2_32.dll" (ByVal s As Long, ByVal hwnd As Long, ByVal wMsg As Long, ByVal lEvent As Long) As Long
Public Declare Function WSAAsyncGetHostByAddr Lib "ws2_32.dll" (ByVal hwnd As Long, ByVal wMsg As Long, ByRef lngAddr As Long, ByVal lngLenght As Long, ByVal lngType As Long, buf As Any, ByVal lngBufLen As Long) As Long
Public Declare Function WSAAsyncGetHostByName Lib "ws2_32.dll" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal strHostName As String, buf As Any, ByVal buflen As Long) As Long
Public Declare Function WSAStartup Lib "ws2_32.dll" (ByVal wVR As Long, lpWSAD As WSAData) As Long
Public Declare Function WSACleanup Lib "ws2_32.dll" () As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function CreateWindowEx Lib "user32" Alias "CreateWindowExA" (ByVal dwExStyle As Long, ByVal lpClassName As String, ByVal lpWindowName As String, ByVal dwStyle As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hWndParent As Long, ByVal hMenu As Long, ByVal hInstance As Long, lpParam As Any) As Long
Private Declare Function DestroyWindow Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function RegisterWindowMessage Lib "user32" Alias "RegisterWindowMessageA" (ByVal lpString As String) As Long
Private Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc&, ByVal hwnd&, ByVal Msg&, ByVal wParam&, ByVal lParam&) As Long
Private Declare Function lstrcpy Lib "kernel32" Alias "lstrcpyA" (ByVal lpString1 As String, ByVal lpString2 As Long) As Long
Private Declare Function lstrlen Lib "kernel32" Alias "lstrlenA" (ByVal lpString As Any) As Long
Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (lpvDest As Any, lpvSource As Any, ByVal cbCopy As Long)
Private Declare Sub RtlMoveMemory Lib "kernel32" (hpvDest As Any, ByVal hpvSource As Long, ByVal cbCopy As Long)
Public Declare Function GlobalAlloc Lib "kernel32" (ByVal wFlags As Long, ByVal dwBytes As Long) As Long
Public Declare Function GlobalFree Lib "kernel32" (ByVal hMem As Long) As Long
Public Declare Function GlobalLock Lib "kernel32" (ByVal hMem As Long) As Long
Public Declare Function GlobalUnlock Lib "kernel32" (ByVal hMem As Long) As Long
Public Declare Function WSAIoctl Lib "ws2_32.dll" (ByVal s As Long, ByVal dwIoControlCode As Long, lpvInBuffer As Any, ByVal cbInBuffer As Long, lpvOutBuffer As Any, ByVal cbOutBuffer As Long, lpcbBytesReturned As Long, lpOverlapped As Long, lpCompletionRoutine As Long) As Long
Private Const GWL_WNDPROC           As Long = -4
Public Const GMEM_FIXED             As Long = &H0
Public Const GMEM_MOVEABLE          As Long = &H2
Public p_lngWindowHandle            As Long
Private m_colSockets                As Collection
Private m_colResolvers              As Collection
Private m_colMemoryBlocks           As Collection
Private m_lngPreviousValue          As Long
Private m_blnGetHostRecv            As Boolean
Private m_blnWinsockInit            As Boolean
Private m_lngMaxMsgSize             As Long
Private Const WM_USER               As Long = &H400
Private m_lngResolveMessage         As Long 'Added: 04-MAR-2002
Public p_lngWinsockMessage          As Long
Private Const OFFSET_4              As Double = 4294967296#
Private Const MAXINT_4              As Double = 2147483647
Private Const OFFSET_2              As Double = 65536
Private Const MAXINT_2              As Double = 32767
Public saZero                       As sockaddr_in
Public WinsockMessage               As Long
Private Function WindowProc(ByVal hwnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    'This the callback function of the window created to hook
    'messages sent by the Winsock service. It handles only two
    'types of messages - network events for the sockets the
    'WSAAsyncSelect fucntion was called for, and the messages
    'sent in response to the WSAAsyncGetHostByName and
    'WSAAsyncGetHostByAddress Winsock API functions.
    '
    'Then the message is received, this function creates illegal
    'reference to the instance of the CSocket class and calls
    'either the PostSocketEvent or PostGetHostEvent method of the
    'class to pass that message to the class.
    Dim objSocket                   As clsSocket 'the illegal reference to an instance of the CSocket class
    Dim lngObjPointer               As Long     'pointer to the existing instance of the CSocket class
    Dim lngEventID                  As Long     'network event
    Dim lngErrorCode                As Long     'code of the error message
    Dim lngMemoryHandle             As Long     'descriptor of the allocated memory object
    Dim lngMemoryPointer            As Long     'pointer to the allocated memory
    Dim lngHostAddress              As Long     '32-bit host address
    Dim strHostName                 As String   'a host hame
    Dim udtHost                     As HOSTENT  'structure of the data in the allocated memory block
    Dim lngIpAddrPtr                As Long     'pointer to the IP address string
    On Error GoTo ERORR_HANDLER
    If uMsg = p_lngWinsockMessage Then
        'All the pointers to the existing instances of the CSocket class
        'are stored in the m_colSockets collection. Key of the collection's
        'item contains a value of the socket handle, and a value of the
        'collection item is the Long value that is a pointer the object,
        'instance of the CSocket class. Since the wParam argument of the
        'callback function contains a value of the socket handle the
        'function has received the network event message for, we can use
        'that value to get the object's pointer. With the pointer value
        'we can create the illegal reference to the object to be able to
        'call any Public or Friend subroutine of that object.
        Set objSocket = SocketObjectFromPointer(CLng(m_colSockets("S" & wParam)))
        'Retrieve the network event ID
        lngEventID = LoWord(lParam)
        'Retrieve the error code
        lngErrorCode = HiWord(lParam)
        'Forward the message to the instance of the CSocket class
        objSocket.PostSocketEvent lngEventID, lngErrorCode
    ElseIf uMsg = m_lngResolveMessage Then
        'A message has been received in response to the call of
        'the WSAAsyncGetHostByName or WSAAsyncGetHostByAddress.
        'Retrieve the error code
        lngErrorCode = HiWord(lParam)
        'The wParam parameter of the callback function contains
        'the task handle returned by the original function call
        '(see the ResolveHost function for more info). This value
        'is used as a key of the m_colResolvers collection item.
        'The item of that collection contains a pointer to the
        'instance of the CSocket class. So, if we know a value
        'of the task handle, we can find out the pointer to the
        'object which called the ResolveHost function in this module.
        'Get the object pointer by the task handle value
        lngObjPointer = CLng(m_colResolvers("R" & wParam))
        'A value of the pointer to the instance of the CSocket class
        'is used also as a key for the m_colMemoryBlocks collection
        'item that contains a handle of the allocated memory block
        'object. That memory block is the buffer where the
        'WSAAsyncGetHostByName and WSAAsyncGetHostByAddress functions
        'store the result HOSTENT structure.
        'Get the handle of the allocated memory block object by the
        'pointer to the instance of the CSocket class.
        lngMemoryHandle = CLng(m_colMemoryBlocks("S" & lngObjPointer))
        'Lock the memory block and get address of the buffer where
        'the HOSTENT structure data is stored.
        lngMemoryPointer = GlobalLock(lngMemoryHandle)
        'Create an illegal reference to the instance of the
        'CSocket class
        Set objSocket = SocketObjectFromPointer(lngObjPointer)
        'Now we can forward the message to that instance.
        If lngErrorCode <> 0 Then
            'If the host was not resolved, pass the error code value
            objSocket.PostGetHostEvent 0, 0, "", lngErrorCode
        Else
            'Move data from the allocated memory block to the
            'HOSTENT structure - udtHost
            CopyMemory udtHost, ByVal lngMemoryPointer, Len(udtHost)
            'Get a 32-bit host address
            CopyMemory lngIpAddrPtr, ByVal udtHost.hAddrList, 4
            CopyMemory lngHostAddress, ByVal lngIpAddrPtr, 4
            'Get a host name
            strHostName = StringFromPointer(udtHost.hName)
            'Call the PostGetHostEvent friend method of the objSocket
            'to forward the retrieved information.
            objSocket.PostGetHostEvent wParam, lngHostAddress, strHostName
        End If
        'The task to resolve the host name is completed, thus we don't
        'need the allocated memory block anymore and corresponding items
        'in the m_colMemoryBlocks and m_colResolvers collections as well.
        'Unlock the memory block
        Call GlobalUnlock(lngMemoryHandle)
        'Free that memory
        Call GlobalFree(lngMemoryHandle)
        'Rremove the items from the collections
        m_colMemoryBlocks.Remove "S" & lngObjPointer
        m_colResolvers.Remove "R" & wParam
        'If there are no more resolving tasks in progress,
        'destroy the collection objects to free resources.
        If m_colResolvers.Count = 0 Then
            Set m_colMemoryBlocks = Nothing
            Set m_colResolvers = Nothing
        End If
    Else
        'Pass other messages to the original window procedure
        WindowProc = CallWindowProc(m_lngPreviousValue, hwnd, uMsg, wParam, lParam)
    End If
EXIT_LABEL:
    Exit Function
ERORR_HANDLER:
    Err.Clear
    'Err.Raise Err.Number, "CSocket.WindowProc", Err.Description
End Function
Public Function RegisterSocket(ByVal lngSocketHandle As Long, ByVal lngObjectPointer As Long) As Boolean
    '********************************************************************************
    'Author    :Oleg Gdalevich
    'Date/Time :17-12-2001
    'Purpose   :Adds the socket to the m_colSockets collection, and
    '           registers that socket with WSAAsyncSelect Winsock API
    '           function to receive network events for the socket.
    '           If this socket is the first one to be registered, the
    '           window and collection will be created in this function as well.
    'Arguments :lngSocketHandle  - the socket handle
    '           lngObjectPointer - pointer to an object, instance of the CSocket class
    'Returns   :If the argument is valid and no error occurred - True.
    '********************************************************************************
    On Error GoTo ERROR_HANDLER
    Dim lngEvents                   As Long
    Dim lngRetValue                 As Long
    If p_lngWindowHandle = 0 Then
        'We have no window to catch the network events.
        'Create a new one.
        p_lngWindowHandle = CreateWinsockMessageWindow
        If p_lngWindowHandle = 0 Then
            'Cannot create a new window.
            'Set the error info to pass to the caller subroutine
            Err.Number = sckOpCanceled
            Err.Description = "The operation was canceled."
            Err.Source = "MSocketSupport.RegisterSocket"
            'Just exit to return False
            Exit Function
        End If
    End If
    'The m_colSockets collection holds information
    'about all the sockets. If the current socket is
    'the first one, create the collection object.
    If m_colSockets Is Nothing Then
        Set m_colSockets = New Collection
        'Debug.Print "The m_colSockets is created"
    End If
    'Add a new item to the m_colSockets collection.
    'The item key contains the socket handle, and the item's data
    'is the pointer to the instance of the CSocket class.
    m_colSockets.Add lngObjectPointer, "S" & lngSocketHandle
    'The lngEvents variable contains a bitmask of events we are
    'going to catch with the window callback function.
    lngEvents = FD_CONNECT Or FD_READ Or FD_WRITE Or FD_CLOSE Or FD_ACCEPT
    'Force the Winsock service to send the network event notifications
    'to the window which handle is p_lngWindowHandle.
    lngRetValue = WSAAsyncSelect(lngSocketHandle, p_lngWindowHandle, p_lngWinsockMessage, lngEvents)    'Modified:04-MAR-2002
    If lngRetValue = SOCKET_ERROR Then
        'If the WSAAsyncSelect call failed this function must
        'return False. In this case, the caller subroutine will
        'raise an error. Let's pass the error info with the Err object.
        RegisterSocket = False
        Err.Number = Err.LastDllError
        Err.Description = GetErrorDescription(Err.LastDllError)
        Err.Source = "MSocketSupport.RegisterSocket"
    Else
        RegisterSocket = True
    End If
    Exit Function
ERROR_HANDLER:
    Err.Clear
    RegisterSocket = False
End Function
Public Function UnregisterSocket(ByVal lngSocketHandle As Long) As Boolean
    '********************************************************************************
    'Author    :Oleg Gdalevich
    'Date/Time :17-12-2001
    'Purpose   :Removes the socket from the m_colSockets collection
    '           If it is the last socket in that collection, the window
    '           and colection will be destroyed as well.
    'Returns   :If the argument is valid and no error occurred - True.
    '********************************************************************************
    If (lngSocketHandle = INVALID_SOCKET) Or (m_colSockets Is Nothing) Then
        'Something wrong with the caller of this function :)
        'Return False
        Exit Function
    End If
    Call WSAAsyncSelect(lngSocketHandle, p_lngWindowHandle, 0&, 0&)
    'Remove the socket from the collection
    m_colSockets.Remove "S" & lngSocketHandle
    UnregisterSocket = True
    If m_colSockets.Count = 0 Then
        'If there are no more sockets in the collection
        'destroy the collection object and the window
        Set m_colSockets = Nothing
        UnregisterSocket = DestroyWinsockMessageWindow
    End If
End Function
Public Function ResolveHost(strHostAddress As String, ByVal lngObjectPointer As Long) As Long
    '********************************************************************************
    'Author    :Oleg Gdalevich
    'Date/Time :17-12-2001
    'Purpose   :Receives requests to resolve a host address from the CSocket class.
    'Returns   :If no errors occurred - ID of the request. Otherwise - 0.
    '********************************************************************************
    'Since this module is supposed to serve several instances of the
    'CSocket class, this function can be called to start several
    'resolving tasks that could be executed simultaneously. To
    'distinguish the resolving tasks the m_colResolvers collection
    'is used. The key of the collection's item contains a pointer to
    'the instance of the CSocket class and the item's data is the
    'Request ID, the value returned by the WSAAsyncGetHostByXXXX
    'Winsock API function. So in order to get the pointer to the
    'instance of the CSocket class by the task ID value the following
    'line of code can be used:
    '
    'lngObjPointer = CLng(m_colResolvers("R" & lngTaskID))
    '
    'The WSAAsyncGetHostByXXXX function needs the buffer (the buf argument)
    'where the data received from DNS server will be stored. We cannot use
    'a local byte array for this purpose as this buffer must be available
    'from another subroutine in this module - WindowProc, also we cannot
    'use a module level array as several tasks can be executed simultaneously
    'At least, we need a dynamic module level array of arrays - too complicated.
    'I decided to use Windows API functions for allocation some memory for
    'each resolving task: GlobalAlloc, GlobalLock, GlobalUnlock, and GlobalFree.
    '
    'To distinguish those memory blocks, the m_colMemoryBlocks collection is
    'used. The key of the collection's item contains value of the object
    'pointer, and the item's value is a handle of the allocated memory
    'block object, value returned by the GlobalAlloc function. So in order to
    'get value of the handle of the allocated memory block object by the
    'pointer to the instance of CSocket class we can use the following code:
    '
    'lngMemoryHandle = CLng(m_colMemoryBlocks("S" & lngObjPointer))
    '
    'Why do we need all this stuff?
    '
    'The problem is that the callback function give us only the resolving task
    'ID value, but we need information about:
    '   - where the data returned from the DNS server is stored
    '   - which instance of the CSocket class we need to post the info to
    '
    'So, if we know the task ID value, we can find out the object pointer:
    '   lngObjPointer = CLng(m_colResolvers("R" & lngTaskID))
    '
    'If we know the object pointer value we can find out where the data is strored:
    '   lngMemoryHandle = CLng(m_colMemoryBlocks("S" & lngObjPointer))
    '
    'That's it. :))
    Dim lngAddress                  As Long '32-bit host address
    Dim lngRequestID                As Long 'value returned by WSAAsyncGetHostByXXX
    Dim lngMemoryHandle             As Long 'handle of the allocated memory block object
    Dim lngMemoryPointer            As Long 'address of the memory block
    Dim strKey                      As String
    'Allocate some memory
    lngMemoryHandle = GlobalAlloc(GMEM_FIXED, MAXGETHOSTSTRUCT)
    If lngMemoryHandle > 0 Then
        'Lock the memory block just to get the address
        'of that memory into the lngMemoryPointer variable
        lngMemoryPointer = GlobalLock(lngMemoryHandle)
        If lngMemoryPointer = 0 Then
            'Memory allocation error
            Call GlobalFree(lngMemoryHandle)
            Exit Function
        Else
            'Unlock the memory block
            GlobalUnlock (lngMemoryHandle)
        End If
    Else
        'Memory allocation error
        Exit Function
    End If
    'If this request is the first one, create the collections
    If m_colResolvers Is Nothing Then
        Set m_colMemoryBlocks = New Collection
        Set m_colResolvers = New Collection
    End If
    strKey = "S" & CStr(lngObjectPointer)
    Call RemoveIfExists(strKey)
    'Remember the memory block location
    m_colMemoryBlocks.Add lngMemoryHandle, strKey
    'Here is a major change. Since version 1.0.6 (08-JULY-2002) the
    'SCocket class doesn't try to resolve the IP address into a
    'domain name while connecting.
    'Try to get 32-bit address
    'lngAddress = inet_addr(strHostAddress)
    lngRequestID = WSAAsyncGetHostByName(p_lngWindowHandle, m_lngResolveMessage, strHostAddress, ByVal lngMemoryPointer, MAXGETHOSTSTRUCT)
    If lngRequestID <> 0 Then
        'If the call of the WSAAsyncGetHostByXXXX is successful, the
        'lngRequestID variable contains the task ID value.
        'Remember it.
        m_colResolvers.Add lngObjectPointer, "R" & CStr(lngRequestID)
        'Return value
        ResolveHost = lngRequestID
    Else
        'If the call of the WSAAsyncGetHostByXXXX is not successful,
        'remove the item from the m_colMemoryBlocks collection.
        m_colMemoryBlocks.Remove ("S" & CStr(lngObjectPointer))
        'Free allocated memory block
        Call GlobalFree(lngMemoryHandle)
        'If there are no more resolving tasks in progress,
        'destroy the collection objects.
        If m_colResolvers.Count = 0 Then
            Set m_colResolvers = Nothing
            Set m_colMemoryBlocks = Nothing
        End If
        'Set the error info.
        Err.Number = Err.LastDllError
        Err.Description = GetErrorDescription(Err.LastDllError)
        Err.Source = "MSocketSupport.ResolveHost"
    End If
End Function
Private Function CreateWinsockMessageWindow() As Long
    '********************************************************************************
    'Author    :Oleg Gdalevich
    'Date/Time :17-12-2001
    'Purpose   :Creates a window to hook the winsock messages
    'Returns   :The window handle
    '********************************************************************************
    'Create a window. It will be used for hooking messages for registered
    'sockets, and we'll not see this window as the ShowWindow is never called.
    p_lngWindowHandle = CreateWindowEx(0&, "STATIC", "SOCKET_WINDOW", 0&, 0&, 0&, 0&, 0&, 0&, 0&, App.hInstance, ByVal 0&)
    If p_lngWindowHandle = 0 Then
        'I really don't know - is this possible? Probably - yes,
        'due the lack of the system resources, for example.
        'In this case the function returns 0.
    Else
        'Register a callback function for the window created a moment ago in this function
        'm_lngPreviousValue - stores the returned value that is the pointer to the previous
        'callback window function. We'll need this value to destroy the window.
        m_lngPreviousValue = SetWindowLong(p_lngWindowHandle, GWL_WNDPROC, AddressOf WindowProc)
        'Just to let the caller know that the function was executed successfully
        CreateWinsockMessageWindow = p_lngWindowHandle
    End If
End Function
Private Function DestroyWinsockMessageWindow() As Boolean
    '********************************************************************************
    'Author    :Oleg Gdalevich
    'Date/Time :17-12-2001
    'Purpose   :Destroyes the window
    'Returns   :If the window was destroyed successfully - True.
    '********************************************************************************
    On Error GoTo ERR_HANDLER
    'Return the previous window procedure
    SetWindowLong p_lngWindowHandle, GWL_WNDPROC, m_lngPreviousValue
    'Destroy the window
    DestroyWindow p_lngWindowHandle
    'Debug.Print "The window " & p_lngWindowHandle & " is destroyed"
    'Reset the window handle variable
    p_lngWindowHandle = 0
    'If no errors occurred, the function returns True
    DestroyWinsockMessageWindow = True
ERR_HANDLER:
    Err.Clear
End Function
Private Function SocketObjectFromPointer(ByVal lngPointer As Long) As clsSocket
    Dim objSocket                   As clsSocket
    CopyMemory objSocket, lngPointer, 4&
    Set SocketObjectFromPointer = objSocket
    CopyMemory objSocket, 0&, 4&
End Function
Private Function LoWord(lngValue As Long) As Long
    LoWord = (lngValue And &HFFFF&)
End Function
Private Function HiWord(lngValue As Long) As Long
    If (lngValue And &H80000000) = &H80000000 Then
        HiWord = ((lngValue And &H7FFF0000) \ &H10000) Or &H8000&
    Else
        HiWord = (lngValue And &HFFFF0000) \ &H10000
    End If
End Function
Public Function GetErrorDescription(ByVal lngErrorCode As Long) As String
    Dim strDesc                     As String
    Select Case lngErrorCode
        Case WSAEACCES
            strDesc = "Permission denied."
        Case WSAEADDRINUSE
            strDesc = "Address already in use."
        Case WSAEADDRNOTAVAIL
            strDesc = "Cannot assign requested address."
        Case WSAEAFNOSUPPORT
            strDesc = "Address family not supported by protocol family."
        Case WSAEALREADY
            strDesc = "Operation already in progress."
        Case WSAECONNABORTED
            strDesc = "Software caused connection abort."
        Case WSAECONNREFUSED
            strDesc = "Connection refused."
        Case WSAECONNRESET
            strDesc = "Connection reset by peer."
        Case WSAEDESTADDRREQ
            strDesc = "Destination address required."
        Case WSAEFAULT
            strDesc = "Bad address."
        Case WSAEHOSTDOWN
            strDesc = "Host is down."
        Case WSAEHOSTUNREACH
            strDesc = "No route to host."
        Case WSAEINPROGRESS
            strDesc = "Operation now in progress."
        Case WSAEINTR
            strDesc = "Interrupted function call."
        Case WSAEINVAL
            strDesc = "Invalid argument."
        Case WSAEISCONN
            strDesc = "Socket is already connected."
        Case WSAEMFILE
            strDesc = "Too many open files."
        Case WSAEMSGSIZE
            strDesc = "Message too long."
        Case WSAENETDOWN
            strDesc = "Network is down."
        Case WSAENETRESET
            strDesc = "Network dropped connection on reset."
        Case WSAENETUNREACH
            strDesc = "Network is unreachable."
        Case WSAENOBUFS
            strDesc = "No buffer space available."
        Case WSAENOPROTOOPT
            strDesc = "Bad protocol option."
        Case WSAENOTCONN
            strDesc = "Socket is not connected."
        Case WSAENOTSOCK
            strDesc = "Socket operation on nonsocket."
        Case WSAEOPNOTSUPP
            strDesc = "Operation not supported."
        Case WSAEPFNOSUPPORT
            strDesc = "Protocol family not supported."
        Case WSAEPROCLIM
            strDesc = "Too many processes."
        Case WSAEPROTONOSUPPORT
            strDesc = "Protocol not supported."
        Case WSAEPROTOTYPE
            strDesc = "Protocol wrong type for socket."
        Case WSAESHUTDOWN
            strDesc = "Cannot send after socket shutdown."
        Case WSAESOCKTNOSUPPORT
            strDesc = "Socket type not supported."
        Case WSAETIMEDOUT
            strDesc = "Connection timed out."
        Case WSATYPE_NOT_FOUND
            strDesc = "Class type not found."
        Case WSAEWOULDBLOCK
            strDesc = "Resource temporarily unavailable."
        Case WSAHOST_NOT_FOUND
            strDesc = "Host not found."
        Case WSANOTINITIALISED
            strDesc = "Successful WSAStartup not yet performed."
        Case WSANO_DATA
            strDesc = "Valid name, no data record of requested type."
        Case WSANO_RECOVERY
            strDesc = "This is a nonrecoverable error."
        Case WSASYSCALLFAILURE
            strDesc = "System call failure."
        Case WSASYSNOTREADY
            strDesc = "Network subsystem is unavailable."
        Case WSATRY_AGAIN
            strDesc = "Nonauthoritative host not found."
        Case WSAVERNOTSUPPORTED
            strDesc = "Winsock.dll version out of range."
        Case WSAEDISCON
            strDesc = "Graceful shutdown in progress."
        Case Else
            strDesc = "Unknown error."
    End Select
    GetErrorDescription = strDesc
End Function
Public Function InitWinsockService() As Long
    'This functon does two things; it initializes the Winsock
    'service and returns value of maximum size of the UDP
    'message. Since this module is supposed to serve multiple
    'instances of the CSocket class, this function can be
    'called several times. But we need to call the WSAStartup
    'Winsock API function only once when the first instance of
    'the CSocket class is created.
    Dim lngRetVal                   As Long     'value returned by WSAStartup
    Dim strErrorMsg                 As String   'error description string
    Dim udtWinsockData              As WSAData  'structure to pass to WSAStartup as an argument
    If Not m_blnWinsockInit Then
        'start up winsock service
        lngRetVal = WSAStartup(&H101, udtWinsockData)
        If lngRetVal <> 0 Then
            'The system cannot load the Winsock library.
            Select Case lngRetVal
                Case WSASYSNOTREADY
                    strErrorMsg = "The underlying network subsystem is not ready for network communication."
                Case WSAVERNOTSUPPORTED
                    strErrorMsg = "The version of Windows Sockets API support requested is not provided by this particular Windows Sockets implementation."
                Case WSAEINVAL
                    strErrorMsg = "The Windows Sockets version specified by the application is not supported by this DLL."
            End Select
            Err.Raise Err.LastDllError, "MSocketSupport.InitWinsockService", strErrorMsg
        Else
            'The Winsock library is loaded successfully.
            m_blnWinsockInit = True
            'This function returns returns value of
            'maximum size of the UDP message
            m_lngMaxMsgSize = IntegerToUnsigned(udtWinsockData.iMaxUdpDg)
            InitWinsockService = m_lngMaxMsgSize
            m_lngResolveMessage = RegisterWindowMessage(App.EXEName & ".ResolveMessage")
            p_lngWinsockMessage = RegisterWindowMessage(App.EXEName & ".WinsockMessage")
        End If
    Else
        'If this function has been called before by another
        'instance of the CSocket class, the code to init the
        'Winsock service must not be executed, but the function
        'returns maximum size of the UDP message anyway.
        InitWinsockService = m_lngMaxMsgSize
    End If
End Function
Public Sub CleanupWinsock()
    '********************************************************************************
    'This subroutine is called from the Class_Terminate() event
    'procedure of any instance of the CSocket class. But the WSACleanup
    'Winsock API function is called only if the calling object is the
    'last instance of the CSocket class within the current process.
    '********************************************************************************
    'If the Winsock library was loaded
    'before and there are no more sockets.
    If m_blnWinsockInit And m_colSockets Is Nothing Then
        'Unload library and free the system resources
        Call WSACleanup
        'Turn off the m_blnWinsockInit flag variable
        m_blnWinsockInit = False
    End If
End Sub
Public Function StringFromPointer(ByVal lPointer As Long) As String
    Dim strTemp                     As String
    Dim lRetVal                     As Long
    'prepare the strTemp buffer
    strTemp = String$(lstrlen(ByVal lPointer), 0)
    'copy the string into the strTemp buffer
    lRetVal = lstrcpy(ByVal strTemp, ByVal lPointer)
    'return a string
    If lRetVal Then StringFromPointer = strTemp
End Function
Public Function UnsignedToLong(Value As Double) As Long
    'The function takes a Double containing a value in the 
    'range of an unsigned Long and returns a Long that you 
    'can pass to an API that requires an unsigned Long
    If Value < 0 Or Value >= OFFSET_4 Then Error 6 ' Overflow
    If Value <= MAXINT_4 Then
        UnsignedToLong = Value
    Else
        UnsignedToLong = Value - OFFSET_4
    End If
End Function
Public Function LongToUnsigned(Value As Long) As Double
    'The function takes an unsigned Long from an API and 
    'converts it to a Double for display or arithmetic purposes
    If Value < 0 Then
        LongToUnsigned = Value + OFFSET_4
    Else
        LongToUnsigned = Value
    End If
End Function
Public Function UnsignedToInteger(Value As Long) As Integer
    'The function takes a Long containing a value in the range 
    'of an unsigned Integer and returns an Integer that you 
    'can pass to an API that requires an unsigned Integer
    If Value < 0 Or Value >= OFFSET_2 Then Error 6 ' Overflow
    If Value <= MAXINT_2 Then
        UnsignedToInteger = Value
    Else
        UnsignedToInteger = Value - OFFSET_2
    End If
End Function
Public Function IntegerToUnsigned(Value As Integer) As Long
    'The function takes an unsigned Integer from and API and 
    'converts it to a Long for display or arithmetic purposes
    If Value < 0 Then
        IntegerToUnsigned = Value + OFFSET_2
    Else
        IntegerToUnsigned = Value
    End If
End Function
Private Sub RemoveIfExists(ByVal strKey As String)
    On Error Resume Next
    m_colMemoryBlocks.Remove strKey
End Sub
Public Function GetHostByNameAlias(ByVal strHostName As String) As Long
    On Error Resume Next
    Dim lpHostent                   As Long
    Dim udtHostent                  As HOSTENT
    Dim AddrList                    As Long
    Dim retIP                       As Long
    retIP = inet_addr(strHostName)
    If retIP = INADDR_NONE Then
        lpHostent = gethostbyname(strHostName)
        If lpHostent <> 0 Then
            CopyMemory udtHostent, ByVal lpHostent, LenB(udtHostent)
            CopyMemory AddrList, ByVal udtHostent.hAddrList, 4
            CopyMemory retIP, ByVal AddrList, udtHostent.hLength
        Else
            retIP = INADDR_NONE
        End If
    End If
    GetHostByNameAlias = retIP
End Function
Public Function GetStrIPFromLong(lngIP As Long) As String
    Dim Bytes(3)                    As Byte
    Call CopyMemory(ByVal VarPtr(Bytes(0)), ByVal VarPtr(lngIP), 4)
    Let GetStrIPFromLong = Bytes(0) & "." & Bytes(1) & "." & Bytes(2) & "." & Bytes(3)
End Function
Public Function EnumNetworkInterfaces() As String
    Dim lngSocketDescriptor         As Long
    Dim lngBytesReturned            As Long
    Dim Buffer                      As INTERFACEINFO
    Dim NumInterfaces               As Integer
    Dim i                           As Integer
    lngSocketDescriptor = socket(AF_INET, SOCK_STREAM, 0)
    If lngSocketDescriptor = INVALID_SOCKET Then
        EnumNetworkInterfaces = INVALID_SOCKET
        Exit Function
    End If
    If WSAIoctl(lngSocketDescriptor, SIO_GET_INTERFACE_LIST, ByVal 0, ByVal 0, Buffer, 1024, lngBytesReturned, ByVal 0, ByVal 0) Then
        EnumNetworkInterfaces = INVALID_SOCKET
        Exit Function
    End If
    NumInterfaces = CInt(lngBytesReturned / 76)
    For i = 0 To NumInterfaces - 1
        EnumNetworkInterfaces = EnumNetworkInterfaces & GetStrIPFromLong(Buffer.iInfo(i).iiAddress.AddressIn.sin_addr) & ";"
    Next i
    CloseSocket lngSocketDescriptor
    EnumNetworkInterfaces = Left$(EnumNetworkInterfaces, Len(EnumNetworkInterfaces) - 1)
End Function
Public Function GetHostNameByAddr(lngIP As Long) As String
    Dim lpHostent                   As Long
    Dim udtHostent                  As HOSTENT
    lpHostent = gethostbyaddr(lngIP, 4, PF_INET)
    If lpHostent = 0 Then Exit Function
    CopyMemory udtHostent, ByVal lpHostent, LenB(udtHostent)
    GetHostNameByAddr = String(256, 0)
    CopyMemory ByVal GetHostNameByAddr, ByVal udtHostent.hName, 256
    GetHostNameByAddr = Left$(GetHostNameByAddr, InStr(1, GetHostNameByAddr, Chr(0)) - 1)
End Function
