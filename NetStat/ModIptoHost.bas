Attribute VB_Name = "ModIptoHost"
Option Explicit
Public Declare Function WSACleanup Lib "ws2_32.dll" () As Long
Public Declare Function WSAGetLastError Lib "ws2_32.dll" () As Long
Private Declare Function WSAStartup Lib "ws2_32.dll" (ByVal wVR As Long, lpWSAD As WSAData) As Long
Private Declare Function gethostbyaddr Lib "ws2_32.dll" (addr As Long, ByVal addr_len As Long, ByVal addr_type As Long) As Long
Private Declare Function gethostbyname Lib "ws2_32.dll" (ByVal host_name As String) As Long
Private Declare Sub RtlMoveMemory Lib "kernel32" (hpvDest As Any, ByVal hpvSource As Long, ByVal cbCopy As Long)
Private Const WSAD_Len              As Long = 256
Private Const WSAS_Len              As Long = 128
Private Type HOSTENT
    hName                           As Long
    hAliases                        As Long
    hAddrType                       As Integer
    hLength                         As Integer
    hAddrList                       As Long
End Type
Private Type WSAData
    wVersion                        As Integer
    wHighVersion                    As Integer
    szDescription(0 To WSAD_Len)    As Byte
    szSystemStatus(0 To WSAS_Len)   As Byte
    iMaxSockets                     As Integer
    iMaxUdpDg                       As Integer
    lpszVendorInfo                  As Long
End Type
Public Function IsIP(ByVal strIP As String) As Boolean
    On Error Resume Next
    Dim t                       As String
    Dim s                       As String
    Dim i                       As Integer
    s = strIP
    While InStr(s, ".") <> 0
        t = Left(s, InStr(s, ".") - 1)
        If IsNumeric(t) And val(t) >= 0 And val(t) <= 255 Then
            s = Mid(s, InStr(s, ".") + 1)
        Else
            Exit Function
        End If
        i = i + 1
    Wend
    t = s
    If IsNumeric(t) And InStr(t, ".") = 0 And Len(t) = Len(Trim(Str(val(t)))) And val(t) >= 0 And val(t) <= 255 And strIP <> "255.255.255.255" And i = 3 Then IsIP = True
    If Err.Number > 0 Then Err.Clear
End Function
Private Function MakeIP(strIP1 As String) As Long
    On Error Resume Next
    Dim strIP As String
    Dim lIP As Long
    strIP = strIP1
    lIP = Left(strIP, InStr(strIP, ".") - 1)
    strIP = Mid(strIP, InStr(strIP, ".") + 1)
    lIP = lIP + Left(strIP, InStr(strIP, ".") - 1) * 256
    strIP = Mid(strIP, InStr(strIP, ".") + 1)
    lIP = lIP + Left(strIP, InStr(strIP, ".") - 1) * 256 * 256
    strIP = Mid(strIP, InStr(strIP, ".") + 1)
    If strIP < 128 Then
        lIP = lIP + strIP * 256 * 256 * 256
    Else
        lIP = lIP + (strIP - 256) * 256 * 256 * 256
    End If
    MakeIP = lIP
    If Err.Number > 0 Then Err.Clear
End Function
Function NameIP(sIP As String) As String
    Dim lIP As Long
    Dim nRet As Long
    Dim hst As HOSTENT
    Dim strHost As String
    Dim strTemp As String
    lIP = MakeIP(sIP)
    nRet = gethostbyaddr(lIP, 4, 2)
    If nRet <> 0 Then
        RtlMoveMemory hst, nRet, Len(hst)
        RtlMoveMemory ByVal strHost, hst.hName, 255
        strTemp = strHost
        strTemp = Left(strTemp, InStr(strTemp, Chr(0)) - 1)
        strTemp = Trim(strTemp)
        NameIP = strTemp
    End If
End Function
Function NameByAddr(strAddr As String) As String
    On Error Resume Next
    If IsConnected = False Then Exit Function
    Dim nRet                        As Long
    Dim lIP                         As Long
    Dim strHost                     As String * 255
    Dim strTemp                     As String
    Dim hst                         As HOSTENT
    If IsIP(strAddr) Then
        lIP = MakeIP(strAddr)
        nRet = gethostbyaddr(lIP, 4, 2)
        If nRet <> 0 Then
            RtlMoveMemory hst, nRet, Len(hst)
            RtlMoveMemory ByVal strHost, hst.hName, 255
            strTemp = strHost
            strTemp = Left(strTemp, InStr(strTemp, Chr(0)) - 1)
            strTemp = Trim(strTemp)
            NameByAddr = strTemp
        Else
            Exit Function
        End If
    Else
        Exit Function
    End If
    If Err.Number > 0 Then Err.Clear
End Function
Function AddrByName(ByVal strHost As String)
    On Error Resume Next
    Dim hostent_addr As Long
    Dim hst As HOSTENT
    Dim hostip_addr As Long
    Dim temp_ip_address() As Byte
    Dim i As Integer
    Dim ip_address As String


    If IsIP(strHost) Then
        AddrByName = strHost
        Exit Function
    End If
    hostent_addr = gethostbyname(strHost)


    If hostent_addr = 0 Then
        MsgBox "Can't resolve hst", , "9001"
        Exit Function
    End If
    RtlMoveMemory hst, hostent_addr, LenB(hst)
    RtlMoveMemory hostip_addr, hst.hAddrList, 4
    ReDim temp_ip_address(1 To hst.hLength)
    RtlMoveMemory temp_ip_address(1), hostip_addr, hst.hLength


    For i = 1 To hst.hLength
        ip_address = ip_address & temp_ip_address(i) & "."
    Next
    ip_address = Mid(ip_address, 1, Len(ip_address) - 1)
    AddrByName = ip_address


    If Err.Number > 0 Then
        MsgBox Err.Description, , Err.Number
        Err.Clear
    End If
End Function
Public Sub IP_Initialize()
    Dim udtWSAData                  As WSAData
    If WSAStartup(257, udtWSAData) Then
        MsgBox Err.Description, , Err.LastDllError
    End If
End Sub
Public Function NameToAddr(ByVal strHost As String)
    Dim ip_list()                   As Byte
    Dim heEntry                     As HOSTENT
    Dim strIPAddr                   As String
    Dim lp_HostEnt                  As Long
    Dim lp_HostIP                   As Long
    Dim iLoop                       As Integer
    On Error GoTo NameToAddrError
    lp_HostEnt = gethostbyname(strHost)
    If lp_HostEnt = 0 Then Exit Function
    RtlMoveMemory heEntry, lp_HostEnt, LenB(heEntry)
    RtlMoveMemory lp_HostIP, heEntry.hAddrList, 4
    ReDim ip_list(1 To heEntry.hLength)
    RtlMoveMemory ip_list(1), lp_HostIP, heEntry.hLength
    For iLoop = 1 To heEntry.hLength
        strIPAddr = strIPAddr & ip_list(iLoop) & "."
    Next
    strIPAddr = Mid(strIPAddr, 1, Len(strIPAddr) - 1)
    NameToAddr = strIPAddr
    Exit Function
NameToAddrError:
    Err.Clear
End Function
