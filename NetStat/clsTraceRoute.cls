VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsTraceRoute"
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
Private Declare Function inet_addr Lib "wsock32" (ByVal s As String) As Long
Private Declare Function inet_ntoa Lib "wsock32.dll" (ByVal addr As Long) As Long
Private Declare Function lstrcpyA Lib "kernel32" (ByVal RetVal As String, ByVal Ptr As Long) As Long
Private Declare Function lstrlenA Lib "kernel32" (ByVal Ptr As Any) As Long
Private Declare Function IcmpCreateFile Lib "icmp.dll" () As Long
Private Declare Function IcmpCloseHandle Lib "icmp.dll" (ByVal IcmpHandle As Long) As Long
Private Declare Function IcmpSendEcho Lib "icmp.dll" (ByVal IcmpHandle As Long, ByVal DestinationAddress As Long, ByVal RequestData As String, ByVal RequestSize As Long, RequestOptions As ICMP_OPTIONS, ReplyBuffer As ICMP_ECHO_REPLY, ByVal ReplySize As Long, ByVal Timeout As Long) As Long
Public Event NewAddress(lHop As Long, sTripTime As String, sAdd As String, sHostName As String)
Private Type ICMP_OPTIONS
    byTTL                           As Byte
    byTOS                           As Byte
    byFlags                         As Byte
    lOptSize                        As Long
    lOptData                        As Long
End Type
Private Type ICMP_ECHO_REPLY
    lAddress                        As Long
    lStat                           As Long
    lRTT                            As Long
    iDataSize                       As Integer
    iRes                            As Integer
    lDataPointer                    As Long
    ICMPOptions                     As ICMP_OPTIONS
    saRetData                       As String * 256
End Type
Private lMaxHopRetry                As Long
Private Const lCharPerPack          As Long = 32 ' This value can be from 32 to 128
Public Function GetIPFromAddress(Address As Long) As String
    Dim lString                     As Long
    lString = inet_ntoa(Address)
    GetIPFromAddress = GetStrFromPtrA(lString)
End Function
Public Function GetStrFromPtrA(ByVal lpszA As Long) As String
    GetStrFromPtrA = String$(lstrlenA(ByVal lpszA), 0)
    Call lstrcpyA(ByVal GetStrFromPtrA, ByVal lpszA)
End Function
Public Function TraceRoute(sIPAddress As String, bHostName As Boolean, lHops As Long, lMaxHopRetry As Long, bMap As Boolean)
    Dim IPOpt                       As ICMP_OPTIONS
    Dim ICMP_Echo                   As ICMP_ECHO_REPLY
    Dim iTTL                        As Integer
    Dim iTTLRetry                   As Integer
    Dim lPort                       As Long
    Dim lNewIP                      As Long
    Dim sAddress                    As String
    Dim sHostIP                     As String
    Dim objNet                      As New clsNetwork
    Dim lMaxHR                      As Long
    Dim udtWSAData                  As WSAData
    lMaxHR = 1
    If WSAStartup(257, udtWSAData) = 0 Then
        ' Check to see if a HostName or IP Address was passed
        If bHostName = True Then
            ' The Host Name was passed, so get the IP Address
            sAddress = objNet.GetIPAdd(sIPAddress)
            sIPAddress = sAddress
        Else
            sAddress = sIPAddress
        End If
        ' Convert the IP Address to a long value
        lNewIP = inet_addr(sAddress)
        ' open an internet file
        lPort = IcmpCreateFile()
        ' Make sure the file was created
        If lPort <> 0 Then
            For iTTL = 1 To lHops
                IPOpt.byTTL = iTTL
                iTTLRetry = SendEcho(lPort, lNewIP, lCharPerPack, sHostIP, ICMP_Echo, IPOpt, bMap)
                If iTTLRetry = 1 Then
                    lMaxHR = lMaxHR + 1
                Else
                    lMaxHR = 1
                End If
                If lMaxHR < lMaxHopRetry Then
                    iTTL = iTTL - iTTLRetry
                End If
                DoEvents
                If sHostIP = sIPAddress Then Exit For
            Next iTTL
            RaiseEvent NewAddress(1, "Trace Complete", "1", "1")
            ' Close the file created
            Call IcmpCloseHandle(lPort)
        Else
            MsgBox "Unable to Open an Icmp File.", vbOKOnly + vbInformation, "Winsock Error"
        End If
        WSACleanup
    Else
        MsgBox "Unable to initialize the Windows Sockets", vbOKOnly + vbInformation, "Socket Error"
    End If
End Function
Private Sub DisplayResults(byTTL As Byte, lTripTime As Long, sHostIP As String)
    Dim sTripTime                   As String
    Dim buff                        As String
    Dim tmp                         As String
    Dim sHostName                   As String
    Dim objNetwork                  As New clsNetwork
    sHostName = objNetwork.GHostName(sHostIP)
    If sHostName = "" Then sHostName = "Not Found"
    Select Case lTripTime
        Case Is < 10
            sTripTime = "<10 ms"
        Case Is > 1200
            sTripTime = "*"
        Case Else
            sTripTime = lTripTime & " ms"
    End Select
    RaiseEvent NewAddress(CStr(byTTL), sTripTime, sHostIP, sHostName)
End Sub
Private Function SendEcho(lPort As Long, lAddress As Long, lCPP As Long, sHostIP As String, ICMP_Echo As ICMP_ECHO_REPLY, ICMP_IPOpt As ICMP_OPTIONS, bMap As Boolean) As Integer
    Dim sData                       As String
    ' Data to send
    sData = String$(lCPP, "a")
    If IcmpSendEcho(lPort, lAddress, sData, Len(sData), ICMP_IPOpt, ICMP_Echo, Len(ICMP_Echo) + 8, 2400) = 1 Then
        sHostIP = GetIPFromAddress(ICMP_Echo.lAddress)
        DisplayResults ICMP_IPOpt.byTTL, ICMP_Echo.lRTT, sHostIP
        SendEcho = 0
    Else
        SendEcho = 1
    End If
End Function

