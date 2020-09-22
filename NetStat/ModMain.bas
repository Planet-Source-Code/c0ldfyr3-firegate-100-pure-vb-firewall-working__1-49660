Attribute VB_Name = "ModIphlpAPI"
Option Explicit
Private Declare Function AllocateAndGetTcpExTableFromStack Lib "iphlpapi.dll" (ByRef pTcpTable As Any, ByVal bOrder As Boolean, ByVal heap As Long, ByVal zero As Long, ByVal Flags As Long) As Long
Private Declare Function AllocateAndGetUdpExTableFromStack Lib "iphlpapi.dll" (ByRef pUdpTable As Any, ByVal bOrder As Boolean, ByVal heap As Long, ByVal zero As Long, ByVal Flags As Long) As Long
Private Declare Function GetProcessHeap Lib "kernel32" () As Long
Private Declare Function GetTcpTable Lib "iphlpapi" (ByRef pTcpTable As Any, ByRef pdwSize As Long, bOrder As Long) As Long
Private Declare Function GetUdpTable Lib "iphlpapi" (ByRef pUdpTable As Any, ByRef pdwSize As Long, bOrder As Long) As Long
Private Declare Function SetTcpEntry Lib "iphlpapi.dll" (ByRef pTcpTable As MIB_TCPROW) As Long
Private Declare Function SetUdpEntry Lib "iphlpapi.dll" (ByRef pTcpTable As MIB_TCPROW) As Long
Private Enum enTcpStates
    TCP_STATE_CLOSED = 1
    TCP_STATE_LISTEN = 2
    TCP_STATE_SYN_SENT = 3
    TCP_STATE_SYN_RCVD = 4
    TCP_STATE_ESTAB = 5
    TCP_STATE_FIN_WAIT1 = 6
    TCP_STATE_FIN_WAIT2 = 7
    TCP_STATE_CLOSE_WAIT = 8
    TCP_STATE_CLOSING = 9
    TCP_STATE_LAST_ACK = 10
    TCP_STATE_TIME_WAIT = 11
    TCP_STATE_DELETE_TCB = 12
End Enum
Private Enum enTcpErrors
    ERROR_BUFFER_OVERFLOW = 111&
    ERROR_NO_DATA = 232&
    ERROR_NOT_SUPPORTED = 50&
    ERROR_SUCCESS = 0&
    ERROR_INVALID_PARAMETER = 87
End Enum
Public Enum enDirection
    Incoming = 1
    Outgoing = 2
End Enum
Public Type MIB_TCPROW             'Type for the GetTcpTable API.
    dwState                         As Long
    dwLocalAddr                     As Long
    dwLocalPort                     As Long
    dwRemoteAddr                    As Long
    dwRemotePort                    As Long
End Type
Private Type MIB_TCPTABLE           'Type for the GetTcpTable API.
    dwNumEntries                    As Long
    Table(1000)                     As MIB_TCPROW
End Type
Private Type MIB_UDPROW             'Type for GetUDPTable API
    dwLocalAddr                     As Long
    dwLocalPort                     As Long
End Type
Private Type MIB_UDPTABLE           'Type for the GetUDPTable API.
    dwNumEntries                    As Long
    Table(1000)                     As MIB_UDPROW
End Type
Private Type MIB_TCPEXROW           'Type for the GetTcpTable API.
    dwState                         As Long     'DWORD   dwState;        // state of the connection
    dwLocalAddr                     As Long     'DWORD   dwLocalAddr;    // address on local computer
    dwLocalPort                     As Long     'DWORD   dwLocalPort;    // port number on local computer
    dwRemoteAddr                    As Long     'DWORD   dwRemoteAddr;   // address on remote computer
    dwRemotePort                    As Long     'DWORD   dwRemotePort;   // port number on remote computer
    dwProcessId                     As Long     'DWORD   dwProcessId;    // Process ID of owner.
End Type
Private Type MIB_UDPEXROW
    dwLocalAddr                     As Long
    dwLocalPort                     As Long
    dwProcessId                     As Long
End Type
Public Type tProcInfo
    lProcID                         As Long
    sPath                           As String
    sUser                           As String
End Type
Public Type tConnectionType         'My type to make it easier to play with the data later on.
    sState                          As String
    sLocalAddr                      As String
    lLocalPort                      As Long
    sRemoteAddr                     As String
    lRemotePort                     As Long
    'lProcessID                      As Long
        'This was tough, if you have time for a little reading then please read on, this is useless info =)
        'When I started this, it was a small project, all I needed was Process ID's and I was using
        'Associative Arrays to get all the info later, this was fine on a small scale.
        'But, with a few connections open here and there, it started to get alot heavier on the loops.
        'So I changed it to add the Path and Owner when it got the Process ID in the first place.
        'Result ? Well, see for yourself, the loops have far less lag and processing power used =)
    ProcInfo                        As tProcInfo
    Direction                       As enDirection
    bTCP                            As Boolean
    Row                             As MIB_TCPROW
End Type
Public Function GetTCPConnections(Optional bShowAll As Boolean = False) As Boolean
    Dim lReturn                     As Long
    Dim lSize                       As Long
    Dim lAddr                       As Long
    Dim lRows                       As Long
    Dim X                           As Long
    Dim rRowXP                      As MIB_TCPEXROW
    Dim rRow                        As MIB_TCPROW
    Dim TcpTable                    As MIB_TCPTABLE
    Dim p_TcpConnections()          As tConnectionType
    LoadProcessUserIDs
    GetProcs
    If g_bXPTable = True Then
        lReturn = AllocateAndGetTcpExTableFromStack(lAddr, True, GetProcessHeap, 2, 2)
            'We pass in a Long to the function, even though C++ would use the actual type structure.
            'This is because C++ uses a memory pointer to the location of the type struct in memory, VB does not, it uses safe arrays, so it fecks this up
            'And ends up returning over 2000000 for the count of items.
            'By doing it this way ,it's gonna receive a pointer to the table *allocated by the function*.
            'Thanks to...
        If lReturn = ERROR_SUCCESS Then 'If succeed...
            CopyMemory lSize, ByVal lAddr, 4 'Get number of entries.
            If bShowAll = True Then ReDim p_TcpConnections(lSize) 'If we are showing them all, might as well redimension the array here.
            For X = 0 To lSize - 1 'Loop through array.
                If cGetInputState(QS_ALLEVENTS) <> 0 Then DoEvents
                If bShowAll = True Then
                    CopyMemory rRowXP, ByVal (lAddr + 4 + X * LenB(rRowXP)), LenB(rRowXP) 'Copy each table individually.
                    'The memory location is calculated by lAddr + 4 + (Size of a Row * Number of rows already done)
                    With p_TcpConnections(X)
                        .lLocalPort = GetTcpPortNumber(rRowXP.dwLocalPort) 'Local Port
                        .ProcInfo.lProcID = rRowXP.dwProcessId 'Process ID
                        .ProcInfo.sPath = Proc_Path(rRowXP.dwProcessId)
                        .ProcInfo.sUser = Proc_UserName(rRowXP.dwProcessId)
                        .lRemotePort = GetTcpPortNumber(rRowXP.dwRemotePort) 'Remote Port
                        .sLocalAddr = GetIpFromLong(rRowXP.dwLocalAddr) 'Local Host Address
                        .sRemoteAddr = GetIpFromLong(rRowXP.dwRemoteAddr) 'Remote Host Address
                        .sState = GetState(rRowXP.dwState) 'State of Connection
                        .Direction = IIf(.lLocalPort = .lRemotePort And .lLocalPort <> 0, enDirection.Incoming, enDirection.Outgoing)
                        .bTCP = True
                    End With
                    With p_TcpConnections(X).Row
                        .dwLocalAddr = rRowXP.dwLocalAddr
                        .dwLocalPort = rRowXP.dwLocalPort
                        .dwRemoteAddr = rRowXP.dwRemoteAddr
                        .dwRemotePort = rRowXP.dwRemotePort
                        .dwState = rRowXP.dwState
                    End With
                Else
                    'Same as in the last one but we only add ones that do not have..
                    'a Local Address of 0.0.0.0 and 127.0.0.1
                    CopyMemory rRowXP, ByVal (lAddr + 4 + X * LenB(rRowXP)), LenB(rRowXP)
                    If Not (GetIpFromLong(rRowXP.dwLocalAddr) = "0.0.0.0" Or GetIpFromLong(rRowXP.dwLocalAddr) = "127.0.0.1") Then
                        lReturn = C_UBound(p_TcpConnections) + 1
                        ReDim Preserve p_TcpConnections(lReturn)
                        With p_TcpConnections(lReturn)
                            .lLocalPort = GetTcpPortNumber(rRowXP.dwLocalPort)
                            .ProcInfo.lProcID = rRowXP.dwProcessId
                            .ProcInfo.sPath = Proc_Path(rRowXP.dwProcessId)
                            .ProcInfo.sUser = Proc_UserName(rRowXP.dwProcessId)
                            .lRemotePort = GetTcpPortNumber(rRowXP.dwRemotePort)
                            .sLocalAddr = GetIpFromLong(rRowXP.dwLocalAddr)
                            .sRemoteAddr = GetIpFromLong(rRowXP.dwRemoteAddr)
                            .sState = GetState(rRowXP.dwState)
                            .Direction = IIf(.lLocalPort = .lRemotePort And .lLocalPort <> 0, enDirection.Incoming, enDirection.Outgoing)
                        End With
                        With p_TcpConnections(lReturn).Row
                            .dwLocalAddr = rRowXP.dwLocalAddr
                            .dwLocalPort = rRowXP.dwLocalPort
                            .dwRemoteAddr = rRowXP.dwRemoteAddr
                            .dwRemotePort = rRowXP.dwRemotePort
                            .dwState = rRowXP.dwState
                        End With
                    End If
                End If
            Next
            g_TcpConnections = p_TcpConnections
            Erase p_TcpConnections
            GetTCPConnections = True
        Else
            GetTCPConnections = False
        End If
    Else
        'This function does the exact same as the last, except it doesn't return ProcessIDs.
        lReturn = GetTcpTable(TcpTable, Len(TcpTable), 0) 'Get the TcpTable (Without Process IDs)
        If lReturn = ERROR_SUCCESS Then 'Sucess ?
            If bShowAll = True Then ReDim p_TcpConnections(TcpTable.dwNumEntries) 'Redimensionalise the array.
            For X = 0 To TcpTable.dwNumEntries 'Loop through the Table.
                If cGetInputState(QS_ALLEVENTS) <> 0 Then DoEvents
                If bShowAll = True Then 'All connections ?
                    With p_TcpConnections(X)
                        .lLocalPort = GetTcpPortNumber(TcpTable.Table(X).dwLocalPort)
                        .ProcInfo.lProcID = 0
                        .lRemotePort = GetTcpPortNumber(TcpTable.Table(X).dwRemotePort)
                        .sLocalAddr = GetIpFromLong(TcpTable.Table(X).dwLocalAddr)
                        .sRemoteAddr = GetIpFromLong(TcpTable.Table(X).dwRemoteAddr)
                        .sState = GetState(TcpTable.Table(X).dwState)
                        .Direction = IIf(.lLocalPort = .lRemotePort And .lLocalPort <> 0, Incoming, Outgoing)
                        .bTCP = True
                        .Row = TcpTable.Table(X)
                    End With
                Else
                    If Not (GetIpFromLong(TcpTable.Table(X).dwLocalAddr) = "0.0.0.0" Or GetIpFromLong(TcpTable.Table(X).dwLocalAddr) = "127.0.0.1") Then
                        lReturn = C_UBound(p_TcpConnections) + 1
                        ReDim Preserve p_TcpConnections(lReturn)
                        With p_TcpConnections(lReturn)
                            .Direction = IIf(.lLocalPort = .lRemotePort And .lLocalPort <> 0, Incoming, Outgoing)
                            .lLocalPort = GetTcpPortNumber(TcpTable.Table(X).dwLocalPort)
                            .ProcInfo.lProcID = 0
                            .lRemotePort = GetTcpPortNumber(TcpTable.Table(X).dwRemotePort)
                            .sLocalAddr = GetIpFromLong(TcpTable.Table(X).dwLocalAddr)
                            .sRemoteAddr = GetIpFromLong(TcpTable.Table(X).dwRemoteAddr)
                            .sState = GetState(TcpTable.Table(X).dwState)
                            .bTCP = True
                            .Direction = IIf(.lLocalPort = .lRemotePort And .lLocalPort <> 0, Incoming, Outgoing)
                            .Row = TcpTable.Table(X)
                        End With
                    End If
                End If
            Next
            g_TcpConnections = p_TcpConnections
            Erase p_TcpConnections
            GetTCPConnections = True
        Else
            GetTCPConnections = False
        End If
    End If
End Function
Public Function GetUDPConnections(Optional bShowAll As Boolean = False) As Boolean
    Dim lReturn                     As Long
    Dim lSize                       As Long
    Dim lAddr                       As Long
    Dim lRows                       As Long
    Dim X                           As Long
    Dim rRowXP                      As MIB_UDPEXROW
    Dim rRow                        As MIB_UDPROW
    Dim TcpTable                    As MIB_UDPTABLE
    Dim p_UdpConnections()          As tConnectionType
    If g_bXPTable = True Then
        lReturn = AllocateAndGetUdpExTableFromStack(lAddr, True, GetProcessHeap, 2, 2)
            'We pass in a Long to the function, even though C++ would use the actual type structure.
            'This is because C++ uses a memory pointer to the location of the type struct in memory, VB does not, it uses safe arrays, so it fecks this up
            'And ends up returning over 2000000 for the count of items.
            'By doing it this way ,it's gonna receive a pointer to the table *allocated by the function*.
            'Thanks to...
        If lReturn = ERROR_SUCCESS Then 'If succeed...
            CopyMemory lSize, ByVal lAddr, 4 'Get number of entries.
            If bShowAll = True Then ReDim p_UdpConnections(lSize) 'If we are showing them all, might as well redimension the array here.
            For X = 0 To lSize - 1 'Loop through array.
                If cGetInputState(QS_ALLEVENTS) <> 0 Then DoEvents
                If bShowAll = True Then
                    CopyMemory rRowXP, ByVal (lAddr + 4 + X * LenB(rRowXP)), LenB(rRowXP) 'Copy each table individually.
                        'The memory location is calculated by lAddr + 4 + (Size of a Row * Number of rows already done)
                    With p_UdpConnections(X)
                        .lLocalPort = GetTcpPortNumber(rRowXP.dwLocalPort)  'Local Port
                        .ProcInfo.lProcID = rRowXP.dwProcessId 'Process ID
                        .ProcInfo.sPath = Proc_Path(rRowXP.dwProcessId)
                        .ProcInfo.sUser = Proc_UserName(rRowXP.dwProcessId)
                        '.lRemotePort = GetTcpPortNumber(rRowXP.dwRemotePort) 'Remote Port
                        .sLocalAddr = GetIpFromLong(rRowXP.dwLocalAddr) 'Local Host Address
                        '.sRemoteAddr = GetIpFromLong(rRowXP.dwRemoteAddr) 'Remote Host Address
                        '.sState = GetState(rRowXP.dwState) 'State of Connection
                        .bTCP = False
                    End With
                Else
                    'Same as in the last one but we only add ones that do not have..
                    'a Local Address of 0.0.0.0 and 127.0.0.1
                    CopyMemory rRowXP, ByVal (lAddr + 4 + X * LenB(rRowXP)), LenB(rRowXP)
                    If Not (GetIpFromLong(rRowXP.dwLocalAddr) = "0.0.0.0" Or GetIpFromLong(rRowXP.dwLocalAddr) = "127.0.0.1") Then
                        lReturn = C_UBound(p_UdpConnections) + 1
                        ReDim Preserve p_UdpConnections(lReturn)
                        With p_UdpConnections(lReturn)
                            .lLocalPort = GetTcpPortNumber(rRowXP.dwLocalPort)
                            '.lRemotePort = GetTcpPortNumber(rRowXP.dwRemotePort)
                            .ProcInfo.lProcID = rRowXP.dwProcessId
                            .ProcInfo.sPath = Proc_Path(rRowXP.dwProcessId)
                            .ProcInfo.sUser = Proc_UserName(rRowXP.dwProcessId)
                            .sLocalAddr = GetIpFromLong(rRowXP.dwLocalAddr)
                            '.sRemoteAddr = GetIpFromLong(rRowXP.dwRemoteAddr)
                            '.sState = GetState(rRowXP.dwState)
                            .bTCP = False
                        End With
                    End If
                End If
            Next
            g_UdpConnections = p_UdpConnections
            Erase p_UdpConnections
            GetUDPConnections = True
        Else
            GetUDPConnections = False
        End If
    Else
        'This function does the exact same as the last, except it doesn't return ProcessIDs.
        lReturn = GetUdpTable(TcpTable, Len(TcpTable), 0) 'Get the TcpTable (Without Process IDs)
        If lReturn = ERROR_SUCCESS Then 'Sucess ?
            If bShowAll = True Then ReDim p_UdpConnections(TcpTable.dwNumEntries) 'Redimensionalise the array.
            For X = 0 To TcpTable.dwNumEntries 'Loop through the Table.
                If cGetInputState(QS_ALLEVENTS) <> 0 Then DoEvents
                If bShowAll = True Then 'All connections ?
                    With p_UdpConnections(X)
                        .lLocalPort = GetTcpPortNumber(TcpTable.Table(X).dwLocalPort)
                        .ProcInfo.lProcID = 0 'We don't get one of these =(
                        '.lRemotePort = GetTcpPortNumber(TcpTable.Table(X).dwRemotePort)
                        .sLocalAddr = GetIpFromLong(TcpTable.Table(X).dwLocalAddr)
                        '.sRemoteAddr = GetIpFromLong(TcpTable.Table(X).dwRemoteAddr)
                        '.sState = GetState(TcpTable.Table(X).dwState)
                    End With
                Else
                    If Not (GetIpFromLong(TcpTable.Table(X).dwLocalAddr) = "0.0.0.0" Or GetIpFromLong(TcpTable.Table(X).dwLocalAddr) = "127.0.0.1") Then
                        lReturn = C_UBound(p_UdpConnections) + 1
                        ReDim Preserve p_UdpConnections(lReturn)
                        With p_UdpConnections(lReturn)
                            .lLocalPort = GetTcpPortNumber(TcpTable.Table(X).dwLocalPort)
                            .ProcInfo.lProcID = 0
                            '.lRemotePort = GetTcpPortNumber(TcpTable.Table(X).dwRemotePort)
                            .sLocalAddr = GetIpFromLong(TcpTable.Table(X).dwLocalAddr)
                            '.sRemoteAddr = GetIpFromLong(TcpTable.Table(X).dwRemoteAddr)
                            '.sState = GetState(TcpTable.Table(X).dwState)
                        End With
                    End If
                End If
            Next
            g_UdpConnections = p_UdpConnections
            Erase p_UdpConnections
            GetUDPConnections = True
        Else
            GetUDPConnections = False
        End If
    End If
End Function
Public Function CloseConnection(Connection As MIB_TCPROW) As Boolean
    'Debug.Print "************************************"
    'Debug.Print ""
    'Debug.Print "Local Address : " & Connection.dwLocalAddr
    'Debug.Print "Local Port : " & Connection.dwLocalPort
    'Debug.Print "Remote Address : " & Connection.dwRemoteAddr
    'Debug.Print "Remote Port : " & Connection.dwRemotePort
    'Debug.Print "State : " & Connection.dwState
    'Debug.Print ""
    'Debug.Print "************************************"
    Connection.dwState = TCP_STATE_DELETE_TCB
    If SetTcpEntry(Connection) = ERROR_SUCCESS Then
        FrmMain.lblStatus.Caption = "Status : Successfully closed connection."
    Else
        FrmMain.lblStatus.Caption = "Status : Could not close connection !"
    End If
End Function
Private Function GetIpFromLong(lngIPAddress As Long) As String
    Dim arrIpParts(3)               As Byte
    CopyMemory arrIpParts(0), lngIPAddress, 4 'This is a pointer to the memory address, so get it !
    GetIpFromLong = CStr(arrIpParts(0)) & "." & CStr(arrIpParts(1)) & "." & CStr(arrIpParts(2)) & "." & CStr(arrIpParts(3))
End Function
Private Function GetTcpPortNumber(DWord As Long) As Long
    GetTcpPortNumber = DWord / 256 + (DWord Mod 256) * 256
End Function
Private Function GetState(lngState As Long) As String
    Select Case lngState
        Case TCP_STATE_CLOSED: GetState = "CLOSED"
        Case TCP_STATE_LISTEN: GetState = "LISTEN"
        Case TCP_STATE_SYN_SENT: GetState = "SYN_SENT"
        Case TCP_STATE_SYN_RCVD: GetState = "SYN_RCVD"
        Case TCP_STATE_ESTAB: GetState = "ESTAB"
        Case TCP_STATE_FIN_WAIT1: GetState = "FIN_WAIT1"
        Case TCP_STATE_FIN_WAIT2: GetState = "FIN_WAIT2"
        Case TCP_STATE_CLOSE_WAIT: GetState = "CLOSE_WAIT"
        Case TCP_STATE_CLOSING: GetState = "CLOSING"
        Case TCP_STATE_LAST_ACK: GetState = "LAST_ACK"
        Case TCP_STATE_TIME_WAIT: GetState = "TIME_WAIT"
        Case TCP_STATE_DELETE_TCB: GetState = "DELETE_TCB"
    End Select
End Function
Public Function C_UBound(aArr() As tConnectionType) As Long
    On Error GoTo ErrClear
    'Nice function to have, instead of trying to handle an error during your functions...
    'Why not just get -1 for an un-initialised array ?
    'As long as you're not using sub 0 lower bound arrays it should be fine.
    C_UBound = UBound(aArr)
    Exit Function
ErrClear:
    C_UBound = -1
    Err.Clear
End Function
