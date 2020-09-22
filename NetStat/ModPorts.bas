Attribute VB_Name = "ModPorts"
Option Explicit
Global g_DBCon                      As ADODB.Connection
Global g_rsPorts                    As ADODB.Recordset
Global g_rsTrojan                   As ADODB.Recordset
Public Enum enPortType
    TCP = "0"
    UDP = "1"
End Enum
Public Function PortDetails(sPortNumber As String, bTrojan As Boolean, pPortType As enPortType) As String
    On Error GoTo ErrClear
    Dim X                           As Integer
    Dim adRs                        As ADODB.Recordset
    Set adRs = New ADODB.Recordset
    PortDetails = ""
    adRs.CursorLocation = adUseClient
    adRs.CursorType = adOpenDynamic
    If bTrojan = True Then
        adRs.Open "SELECT DISTINCT fldTrojanName FROM tblTrojanPorts WHERE fldPort = " & sPortNumber & " AND fldType = '" & IIf(pPortType = TCP, "TCP", "UDP") & "' ORDER BY fldRegisterName", g_DBCon, adOpenDynamic, adLockReadOnly
    Else
        adRs.Open "SELECT DISTINCT fldRegisterName FROM tblRegisteredPorts WHERE fldPort = " & sPortNumber & " AND fldType = '" & IIf(pPortType = TCP, "TCP", "UDP") & "' ORDER BY fldRegisterName", g_DBCon, adOpenDynamic, adLockReadOnly
    End If
    If Not adRs.EOF Then
        adRs.MoveLast
        adRs.MoveFirst
    End If
    'PortDetails = adRs.Fields("fldTrojanName")
    If bTrojan = True Then
        PortDetails = FirstUCase(adRs.Fields("fldTrojanName"))
    Else
        PortDetails = FirstUCase(adRs.Fields("fldRegisterName"))
    End If
ErrClear:
    Err.Clear
    'adRs.Close
    Set adRs = Nothing
End Function
Private Function FirstUCase(sText As String)
    If Len(sText) > 1 Then
        FirstUCase = UCase(Mid(sText, 1, 1)) & Mid(sText, 2)
    Else
        FirstUCase = UCase(sText)
    End If
End Function
Public Function MakeADOConnection() As ADODB.Connection
    Dim adCon                       As ADODB.Connection
    Dim strCon                      As String
    On Error GoTo Error_MakeConnection
    Set adCon = New ADODB.Connection
    adCon.CursorLocation = adUseClient
    strCon = GetConnectionString()
    adCon.Open strCon
    Set MakeADOConnection = adCon
    Exit Function
Error_MakeConnection:
    MsgBox Err.Number & ":" & Err.Description
End Function
Private Function GetConnectionString() As String
    GetConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & FixPath(App.Path) & "Ports.mdb"
End Function
