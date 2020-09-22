VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsIPInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private Declare Function lstrcpy Lib "kernel32" Alias "lstrcpyA" (ByVal lpString1 As String, ByVal lpString2 As Long) As Long
Private Declare Function lstrlen Lib "kernel32" Alias "lstrlenA" (ByVal lpString As Any) As Long
Private Declare Sub RtlMoveMemory Lib "kernel32" (hpvDest As Any, ByVal hpvSource As Long, ByVal cbCopy As Long)
Private Type SERVENT
    s_name                          As Long
    s_aliases                       As Long
    s_port                          As Integer
    s_proto                         As Long
End Type
Private Const OFFSET_2              As Double = 65536
Private WithEvents cSox             As clsSocket
Attribute cSox.VB_VarHelpID = -1
Private sCurrentIP                  As String
Private sServer                     As String
Private lPort                       As Long
Private sFinal                      As String
Private bRetry                      As Boolean
Event Finished(sCompany As String, sCountry As String, sIP As String)
Private Sub Class_Initialize()
    IP_Initialize
    Set cSox = New clsSocket
End Sub
Private Sub Class_Terminate()
    WSACleanup
    Set cSox = Nothing
End Sub
Public Sub GetDetails(sIP As String)
    lPort = ServicePort("whois")
    sServer = "whois.arin.net"
    sCurrentIP = sIP
    sFinal = ""
    cSox.CloseSocket
    cSox.Connect sServer, lPort
End Sub
Private Sub cSox_OnClose()
    ParseInfo sFinal
End Sub
Private Sub cSox_OnConnect()
    sFinal = ""
    cSox.SendData sCurrentIP & vbCrLf
End Sub
Private Sub cSox_OnDataArrival(ByVal bytesTotal As Long)
    Dim sData                       As String
    Call cSox.GetData(sData, vbString)
    DataArrival (sData)
End Sub
Private Sub DataArrival(sData As String)
    Dim sServ                   As String
    Dim sArr()                  As String
    Dim X                       As Long
    sArr = Split(sData, vbLf)
    For X = 0 To UBound(sArr)
        If Mid(sArr(X), 1, 15) = "ReferralServer:" Then
            sServ = Mid(sArr(X), 25)
            sFinal = ""
            sServer = sServ
            bRetry = True
            Call cSox.CloseSocket
            Set cSox = Nothing
            Set cSox = New clsSocket
            Call cSox.Connect(sServer, lPort)
            Exit Sub
        End If
    Next
    If MatchSpec(sData, ".*LACNIC.*") = True Then
        sServ = "whois.apnic.net"
    ElseIf MatchSpec(sData, ".*APNIC.*") = True Then
        sServ = "whois.apnic.net"
    ElseIf MatchSpec(sData, ".*APNIC-.*") = True Then
        sServ = "whois.aunic.net"
    ElseIf MatchSpec(sData, ".*AUNIC-AU.*") = True Then
        sServ = "whois.ripe.net"
    ElseIf MatchSpec(sData, ".*NETBLK-RIPE.*") = True Then
        sServ = "whois.ripe.net"
    ElseIf MatchSpec(sData, ".*NET-RIPE.*") = True Then
        sServ = "whois.ripe.net"
    ElseIf MatchSpec(sData, ".*-RIPE.*") = True Then
        sServ = "whois.ripe.net"
    ElseIf MatchSpec(sData, ".*RIPE-.*") = True Then
        sServ = "whois.ripe.net"
    ElseIf MatchSpec(sData, ".*NETBLK-BRAZIL.*") = True Then
        sServ = "whois.nic.br"
    ElseIf MatchSpec(sData, ".*whois.nic.ad.jp.*") = True Then
        sServ = "whois.nic.ad.jp"
    ElseIf MatchSpec(sData, ".*whois.telstra.*") = True Then
        sServ = "whois.telstra.net"
    ElseIf MatchSpec(sData, ".*whois.exodus.*") = True Then
        sServ = "rwhois.exodus.net"
        lPort = 4321
    ElseIf MatchSpec(sData, ".*rwhois.verio.*") = True Then
        sServ = "rwhois.verio.net"
        lPort = 4321
    ElseIf MatchSpec(sData, ".*rwhois.dnai.*") = True Then
        sServ = "rwhois.dnai.com"
        lPort = 4321
    ElseIf MatchSpec(sData, ".*rwhois.digex.*") = True Then
        sServ = "rwhois.digex.net"
        lPort = 4321
    ElseIf MatchSpec(sData, ".*rwhois.internex.*") = True Then
        sServ = "rwhois.internex.net"
        lPort = 4321
    ElseIf MatchSpec(sData, ".rwhois.concentric.*") = True Then
        sServ = "rwhois.concentric.net"
        lPort = 4321
    ElseIf MatchSpec(sData, ".*rwhois.oar.*") = True Then
        sServ = "rwhois.oar.net"
        lPort = 4321
    ElseIf MatchSpec(sData, ".*rwhois.elan.") = True Then
        sServ = "rwhois.elan.net"
        lPort = 4321
    ElseIf MatchSpec(sData, ".*rwhois.cais.*") = True Then
        sServ = "rwhois.cais.net"
        lPort = 4321
    ElseIf MatchSpec(sData, ".*rwhois.cogentco.*") = True Then
        sServ = "rwhois.cogentco.com"
        lPort = 4321
    ElseIf MatchSpec(sData, ".*rwhois.beanfield.*") = True Then
        sServ = "rwhois.beanfield.net"
        lPort = 4321
    ElseIf MatchSpec(sData, ".*JPNIC*.*") = True Then
        sServ = "whois.nic.ad.jp"
    ElseIf MatchSpec(sData, ".*JNIC*.*") = True Then
        sServ = "whois.nic.ad.jp"
    ElseIf MatchSpec(sData, ".*whois.nic.or.kr.*") = True Then
        sServ = "whois.nic.or.kr"
    End If
    If Len(sServ) > 0 Then
        sFinal = ""
        sServer = sServ
        bRetry = True
        Call cSox.CloseSocket
        Call cSox.Connect(sServer, lPort)
        Exit Sub
    End If
    sFinal = sFinal & sData
End Sub
Private Function EndOfWhiteSpace(sText As String, iPos As Integer) As Integer
    Dim X                           As Integer
    For X = iPos To Len(sText)
        If Not Mid(sText, X, 1) = " " Then
            EndOfWhiteSpace = X
            Exit Function
        End If
    Next
    EndOfWhiteSpace = 0
End Function
Private Sub ParseInfo(sData As String)
    If Not cSox.State = sckClosed Then Exit Sub
    Dim sOrg                    As String
    Dim sCountry                As String
    Dim sCustomer               As String
    Dim X                       As Integer
    Dim sArr()                  As String
    sArr = Split(sData, vbLf)
    X = UBound(sArr)
    For X = 0 To UBound(sArr)
        If LCase(Mid(sArr(X), 1, 6)) = "descr:" Then
            If Len(sCustomer) > 0 Then sCustomer = sCustomer & ", "
            sCustomer = sCustomer & Mid(sArr(X), EndOfWhiteSpace(sArr(X), 7))
        ElseIf LCase(Mid(sArr(X), 1, 8)) = "country:" Then
            sCountry = Mid(sArr(X), EndOfWhiteSpace(sArr(X), 9))
        ElseIf Mid(sArr(X), 1, 8) = "OrgName:" Then
            sOrg = Mid(sArr(X), EndOfWhiteSpace(sArr(X), 9))
        ElseIf Mid(sArr(X), 1, 9) = "CustName:" Then
            sCustomer = Mid(sArr(X), EndOfWhiteSpace(sArr(X), 10))
        End If
    Next
    If Len(sCountry) = 0 And Not Mid(sCurrentIP, 1, 1) = "+" Then
        sServer = "whois.arin.net"
        sCurrentIP = "+" & sCurrentIP
        lPort = ServicePort("whois")
        bRetry = True
        Call cSox.CloseSocket
        Call cSox.Connect(sServer, lPort)
    ElseIf Len(sCountry) = 0 Then
        RaiseEvent Finished("", "", sCurrentIP)
    Else
        If Len(sCustomer) > 0 Then
            RaiseEvent Finished(sCustomer, sCountry, sCurrentIP)
        Else
            RaiseEvent Finished(sOrg, sCountry, sCurrentIP)
        End If
    End If
End Sub
Private Function ServicePort(sService As String) As Long
    Dim udt_servent                 As SERVENT
    Dim lngPointer                  As Long
    Dim colAliases                  As Collection
    Dim varAlias                    As Variant
    lngPointer = getservbyname(Trim(sService), "tcp")
    If lngPointer <> 0 Then
        RtlMoveMemory udt_servent, lngPointer, LenB(udt_servent)
            If Len(udt_servent.s_name) > 0 Then
                'Debug.Print "Name:" & StringFromPointer(udt_servent.s_name)
                'Debug.Print "Proto:" & StringFromPointer(udt_servent.s_proto)
                'Debug.Print "Port:" & IntegerToUnsigned(ntohs(udt_servent.s_port))
                ServicePort = IntegerToUnsigned(ntohs(udt_servent.s_port))
                Set colAliases = PtrArrayToStrCollection(udt_servent.s_aliases)
                If Not colAliases Is Nothing Then
                    For Each varAlias In colAliases
                        'Debug.Print "Alias:" & varAlias
                    Next
                End If
            End If
    End If
End Function
Private Function StringFromPointer(ByVal lPointer As Long) As String
    Dim strTemp                     As String
    Dim lRetVal                     As Long
    'prepare the strTemp buffer
    strTemp = String$(lstrlen(ByVal lPointer), 0)
    'copy the string into the strTemp buffer
    lRetVal = lstrcpy(ByVal strTemp, ByVal lPointer)
    'return a string
    If lRetVal Then StringFromPointer = strTemp
End Function
Private Function IntegerToUnsigned(Value As Integer) As Long
    'The function takes an unsigned Integer from and API and 
    'converts it to a Long for display or arithmetic purposes
    If Value < 0 Then
        IntegerToUnsigned = Value + OFFSET_2
    Else
        IntegerToUnsigned = Value
    End If
End Function
Private Function PtrArrayToStrCollection(lngPointer As Long) As Collection
    Dim lngTempPointer              As Long
    Dim colList                     As Collection
    If lngPointer > 0 Then
        Set colList = New Collection
        'Get poointer to the first array element
        RtlMoveMemory lngTempPointer, lngPointer, 4
        Do Until lngTempPointer = 0
            colList.Add StringFromPointer(lngTempPointer)
            'Get pointer to the next array element
            lngPointer = lngPointer + 4
            RtlMoveMemory lngTempPointer, lngPointer, 4
        Loop
        Set PtrArrayToStrCollection = colList
    End If
End Function

