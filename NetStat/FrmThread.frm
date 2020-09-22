VERSION 5.00
Begin VB.Form FrmThread 
   Caption         =   "Form1"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer tmrThread 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   720
      Top             =   1440
   End
End
Attribute VB_Name = "FrmThread"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public Sub DoThread()
    Call tmrThread_Timer
End Sub
Private Sub Form_Load()
    Me.Hide
    Height = 1
    Width = 1
    Top = Screen.Height * 2
    Left = Screen.Width * 2
    Debug.Print "hidden"
End Sub
Private Sub tmrThread_Timer()
    Dim X                           As Long
    Dim i                           As Long
    Dim Item                        As ListItem
    Dim bComplete                   As Boolean
    Dim sKey                        As String
    Dim bFound                      As Boolean
    Dim bAnother                    As Boolean
    Dim aArray()                    As tConnectionType
    Dim iReturn                     As tPicIndex
    Dim sDisplay                    As tChecking
    Dim List                        As ListView
    Dim iCount                      As Integer
    Dim Str()                       As String
    Dim sTmp                        As String
    Static iIncoming(3)             As Integer
    Static iOutgoing(3)             As Integer
    Static iBlocked(3)              As Integer
    Static bOnline                  As Boolean
    If g_bStopAll = True Then g_bWorking = False: Exit Sub
    If g_bWorking = True Then Exit Sub
    g_bWorking = True
    'There was some weird stuff happening a while ago, this seemed to fix it.
    'It seems impossible but true.
    'I think the function *was* taking more time to work than the timer interval so it was re-itterating before it finished the last tick.
    If g_bStopAll = True Then g_bWorking = False: Exit Sub
    If IsConnected = True And bOnline = False Then
        FrmMain.cmbInterface.Clear
        Str = Split(EnumNetworkInterfaces(), ";")
        For X = 0 To UBound(Str)
            If Str(X) <> "127.0.0.1" Then
                FrmMain.cmbInterface.AddItem Str(X) & " [" & GetHostNameByAddr(inet_addr(Str(X))) & "]"
            End If
        Next
        If FrmMain.cmbInterface.ListCount > 0 Then FrmMain.cmbInterface.ListIndex = 0
    ElseIf IsConnected = False And bOnline = True Then
        FrmMain.cmbInterface.Clear
    End If
    bOnline = IsConnected
ReStart:
    iBlocked(1 + iCount) = iBlocked(0 + iCount)
    iBlocked(0 + iCount) = 0
    iIncoming(1 + iCount) = iIncoming(0 + iCount)
    iIncoming(0 + iCount) = 0
    iOutgoing(1 + iCount) = iOutgoing(0 + iCount)
    iOutgoing(0 + iCount) = 0
    If iCount = 0 Then
        Set List = FrmMain.lstTCPConnections
    Else
        Set List = FrmMain.lstUDPConnections
    End If
    GetProcs 'Get all Process's into our Associative Array
    LoadProcessUserIDs
    Erase aArray
    If iCount = 0 Then  'If TCP Table
        bComplete = GetTCPConnections
        aArray() = g_TcpConnections           'Set the working array.
    Else                                    'If UDP Table
        bComplete = GetUDPConnections
        aArray() = g_UdpConnections           'Set the working array.
    End If
    If g_bStopAll = True Then g_bWorking = False: Exit Sub
    If bComplete Then 'If it succeeded.
        'The reason I went about this the long way was simple.
        'Sure I could have removed all entries, and then added them again, even with LockWindow...
        'It still flickered if only once, it was still flickering and would remove a selected item.
        'So, what I do is, check for the existance of every single connection, and then check which ones aren't
        'There anymore, and remove/add as needed.
        'It does take a lil longer than just adding them all every time, but its worth it.
        'Hey, it works =) And its flickerless =D
        For i = 1 To List.ListItems.Count
            DoEvents
            List.ListItems(i).Tag = "" 'Set the tag to nothing so that we can check these on the fly.
        Next
        For X = 0 To C_UBound(aArray) 'Loop through the connection array.
            DoEvents
            If g_bStopAll = True Then g_bWorking = False: Exit Sub
            If aArray(X).Direction = enDirection.Incoming Then
                iIncoming(0 + iCount) = iIncoming(0 + iCount) + 1
            Else
                iOutgoing(0 + iCount) = iOutgoing(0 + iCount) + 1
            End If
            If iCount = 0 Then  'If TCP
                sKey = aArray(X).sLocalAddr & ":" & CStr(aArray(X).lLocalPort) & ":" & aArray(X).sRemoteAddr & CStr(aArray(X).lRemotePort)
            Else                                    'If UDP
                sKey = aArray(X).sLocalAddr & ":" & CStr(aArray(X).lLocalPort) 'UDP Only has local port and addr, so we have to restructure our key.
            End If
            bFound = False
            For i = 1 To List.ListItems.Count
                DoEvents
                If List.ListItems(i).key = sKey Then
                    List.ListItems(i).Tag = "-" 'This connection and item exists, so we mark it so we know we have counted it.
                    If iCount = 0 Then
                        If Not StrComp(List.ListItems(i).ListSubItems(5).Text, aArray(X).sState) = 0 Then List.ListItems(i).ListSubItems(5).Text = aArray(X).sState
                    End If
                    bFound = True
                    If List.ListItems(i).Checked = True Then
                        List.ListItems(i).Checked = False
                        Call CloseConnection(aArray(X).Row)
                    End If
                End If
            Next i
            If bFound = False Then 'If the connection was not allready in our list.
                With aArray(X)
                    If g_bStopAll = True Then g_bWorking = False: Exit Sub
                    DoEvents
                    sDisplay = CheckConnection(aArray(X))
                    If sDisplay.bBlocked = True Then iBlocked(0 + iCount) = iBlocked(0 + iCount) + 1
                    If g_bXPTable = True Then
                        If Len(.ProcInfo.sPath) = 0 Then
                            Set Item = List.ListItems.Add(, sKey, "[System Process]") 'Add the item
                        Else
                            Set Item = List.ListItems.Add(, sKey, sDisplay.sName)  'Add the item
                        End If
                        Item.ListSubItems.Add(, , CStr(.sLocalAddr)).Tag = Time$
                    Else
                        Set Item = List.ListItems.Add(, sKey, CStr(.sLocalAddr)) 'Add the item
                    End If
                    Item.Checked = False
                    Item.Tag = "-" 'Set the tag for deletion.
                    Item.ListSubItems.Add , , CStr(.lLocalPort)
                    If iCount = 0 Then 'If TCP
                        If FrmMain.chkName.Value = vbUnchecked Then
                            sTmp = NameByAddr(.sRemoteAddr)
                            Item.ListSubItems.Add(, , IIf(Len(sTmp) > 0, sTmp, .sRemoteAddr)).Tag = .sRemoteAddr
                        Else
                            Item.ListSubItems.Add(, , .sRemoteAddr).Tag = .sRemoteAddr
                        End If
                        Item.ListSubItems.Add , , CStr(.lRemotePort)
                        Item.ListSubItems.Add , , .sState
                    End If
                    If g_bXPTable = True Then
                        If .ProcInfo.lProcID > 0 And Len(.ProcInfo.sPath) > 0 Then
                            If FrmMain.chkIcons.Value = vbChecked Then
                                iReturn = GetIcon(.ProcInfo.sPath, Item.Index, g_sShell32Path, FrmMain.Pic16, FrmMain.Pic32, FrmMain.ImgL32, FrmMain.ImgL16)
                                List.Icons = FrmMain.ImgL32
                                List.SmallIcons = FrmMain.ImgL16
                                Item.Icon = iReturn.Pic32
                                Item.SmallIcon = iReturn.Pic16
                            End If
                        End If
                        Item.ListSubItems.Add , , .ProcInfo.sUser
                        Item.ListSubItems(2).Tag = LCase(.ProcInfo.sPath)
                        Item.ListSubItems.Add , , .ProcInfo.lProcID    'Only if we are on XP do we get Process ID's
                    End If
                    If sDisplay.bBlocked = True Then
                        If iCount = 0 Then
                            If Len(sDisplay.sName) > 0 Then FrmBar.ShowAlert ("Blocked " & sDisplay.sName & IIf(aArray(X).Direction = Incoming, " from accepting connection from " & Item.ListSubItems(3).Text, " from connecting to " & Item.ListSubItems(3).Text))
                        Else
                            If Len(sDisplay.sName) > 0 Then FrmBar.ShowAlert ("Blocked " & sDisplay.sName & " on UDP Port " & Item.ListSubItems(2).Text)
                        End If
                    Else
                        If iCount = 0 Then
                            If Len(sDisplay.sName) > 0 Then FrmBar.ShowAlert ("Allowed " & sDisplay.sName & IIf(aArray(X).Direction = Incoming, " to accepting connection from " & Item.ListSubItems(3).Text, " to connect to " & Item.ListSubItems(3).Text))
                        Else
                            If Len(sDisplay.sName) > 0 Then FrmBar.ShowAlert ("Allowed " & sDisplay.sName & " on UDP Port " & Item.ListSubItems(2).Text)
                        End If
                    End If
                End With
            End If
            'The above line will speed up this little timer quite a bit.
            'What it does is check to see are there any pending messages this window needs to process.
            'If their is, it calls a DoEvents so that the window can catch up on itself instead of locking up.
        Next X
Starter:
        If g_bStopAll = True Then g_bWorking = False: Exit Sub
        For i = 1 To List.ListItems.Count
            DoEvents
            If Len(List.ListItems(i).Tag) = 0 Then 'Remove all the items that we did not mark.
                bAnother = False
                For X = 1 To FrmMain.lstTCPConnections.ListItems.Count
                    DoEvents
                    If Not X = i And FrmMain.lstTCPConnections.ListItems(X).ListSubItems(2).Tag = List.ListItems(i).ListSubItems(2).Tag And Not List.ToolTipText = FrmMain.lstUDPConnections.ToolTipText Then
                        bAnother = True
                    End If
                Next
                For X = 1 To FrmMain.lstUDPConnections.ListItems.Count
                    DoEvents
                    If Not X = i And FrmMain.lstUDPConnections.ListItems(X).ListSubItems(2).Tag = List.ListItems(i).ListSubItems(2).Tag And Not List.ToolTipText = FrmMain.lstTCPConnections.ToolTipText Then
                        bAnother = True
                    End If
                Next
                If Len(List.ListItems(i).ListSubItems(2).Tag) > 0 And bAnother = False Then If FindPrograms(List.ListItems(i).ListSubItems(2).Tag) > -1 Then FrmMain.lstPrograms.ListItems(FindPrograms(List.ListItems(i).ListSubItems(2).Tag)).SmallIcon = 13
                List.ListItems.Remove (i)
                GoTo Starter
            End If
        Next
        If FrmMain.lstTCPConnections.ListItems.Count = 0 And FrmMain.lstUDPConnections.ListItems.Count = 0 And (FrmMain.ImgL16.ListImages.Count > 0 Or FrmMain.ImgL32.ListImages.Count > 0) Then
            FrmMain.lstTCPConnections.Icons = Nothing
            FrmMain.lstTCPConnections.SmallIcons = Nothing
            FrmMain.lstUDPConnections.Icons = Nothing
            FrmMain.lstUDPConnections.SmallIcons = Nothing
            FrmMain.ImgL16.ListImages.Clear
            FrmMain.ImgL32.ListImages.Clear
        End If
    End If
    If g_bStopAll = True Then g_bWorking = False: Exit Sub
    DoEvents
    Call FrmMain.CheckTraffic(iIncoming, iOutgoing, iBlocked, iCount)
    If iCount = 0 Then
        iCount = 2
        GoTo ReStart
    End If
    Set Item = Nothing
    Set List = Nothing
    g_bWorking = False
End Sub
Private Function CheckConnection(aConnection As tConnectionType) As tChecking
    Dim X                           As Integer
    Dim sTmp                        As String
    Dim iFound                      As Integer
    Dim sName                       As String
    Dim iTmp                        As Integer
    Dim sShortTmp                   As String
    iFound = -1
    sName = FileInfo(aConnection.ProcInfo.sPath, [File Description])  'Get File Description from Executable
    If Len(sName) = 0 Then 'If it doesn't have a File Description...
        iTmp = InStrRev(aConnection.ProcInfo.sPath, "\") 'Find the last \
        If iTmp > 0 Then
            sName = Mid(aConnection.ProcInfo.sPath, iTmp + 1) 'Take from \ to the end
        Else
            sName = aConnection.ProcInfo.sPath 'If no \, the whole file name, this should never happen I don't think.
        End If
    End If
    If aConnection.bTCP = True Then 'If its a TCP connection...
        For X = 1 To FrmMain.lstIPs.ListItems.Count 'Loop through list of IP's to block.
            If cGetInputState(QS_ALLEVENTS) <> 0 Then DoEvents
            If MatchSpec(aConnection.sRemoteAddr, FrmMain.lstIPs.ListItems(X).Text) = True Then
                'Using a windows API, we check Regular Expression matching.
                'So 127.0.0.* will work or, *.castlrea.eircom.net etc
                If aConnection.Direction = Incoming And (FrmMain.lstIPs.ListItems(X).ListSubItems(1).Tag = "0" Or FrmMain.lstIPs.ListItems(X).ListSubItems(1).Tag = "1") Then
                    'If Outgoing and we want to block In or Both
                    CloseConnection aConnection.Row 'Close connection using the Row table.
                    KillProcess aConnection.ProcInfo.lProcID  'Kill process using Process ID
                    CheckConnection.sName = sName 'Return exe name or description
                    CheckConnection.bBlocked = True 'Return blocked status.
                    Exit Function
                ElseIf aConnection.Direction = Outgoing And (FrmMain.lstIPs.ListItems(X).ListSubItems(1).Tag = "0" Or FrmMain.lstIPs.ListItems(X).ListSubItems(1).Tag = "2") Then
                    'If Outgoing and we want to block Out or Both
                    CloseConnection aConnection.Row 'Close connection using the Row table.
                    KillProcess aConnection.ProcInfo.lProcID  'Kill process using Process ID
                    CheckConnection.sName = sName 'Return exe name or description
                    CheckConnection.bBlocked = True 'Return blocked status.
                    Exit Function
                End If
            End If
        Next
        For X = 1 To FrmMain.lstPorts.ListItems.Count 'Loop through Ports
            If aConnection.Direction = Incoming Then 'If incoming.
                If FrmMain.lstPorts.ListItems(X).Text = aConnection.lLocalPort Then 'If the port is the local port, on incoming connections both ports are the same anyway.
                    If (FrmMain.lstPorts.ListItems(X).ListSubItems(1).Tag = "0" Or FrmMain.lstPorts.ListItems(X).ListSubItems(1).Tag = "1") Then 'If we want to block In or Both
                        CloseConnection aConnection.Row
                        KillProcess aConnection.ProcInfo.lProcID
                        CheckConnection.sName = sName
                        CheckConnection.bBlocked = True
                        Exit Function
                    End If
                    Exit For
                End If
            Else
                If FrmMain.lstPorts.ListItems(X).Text = aConnection.lLocalPort Or FrmMain.lstPorts.ListItems(X).Text = aConnection.lRemotePort Then 'If local or remote is the port we wanna block.
                    If (FrmMain.lstPorts.ListItems(X).ListSubItems(1).Tag = "0" Or FrmMain.lstPorts.ListItems(X).ListSubItems(1).Tag = "2") Then 'If we want to block Out or Both
                        CloseConnection aConnection.Row
                        KillProcess aConnection.ProcInfo.lProcID
                        CheckConnection.sName = sName
                        CheckConnection.bBlocked = True
                        Exit Function
                    End If
                    Exit For
                End If
            End If
        Next
    End If
    sTmp = aConnection.ProcInfo.sPath  'Get the Process Name
    sShortTmp = GetShortPath(sTmp)
    If Len(sTmp) > 0 Then
        For X = 0 To T_UBound(g_aPrograms) 'Loop through the program names.
            If cGetInputState(QS_ALLEVENTS) <> 0 Then DoEvents
            If StrComp(sTmp, g_aPrograms(X).sLocation, vbTextCompare) = 0 Or StrComp(sShortTmp, g_aPrograms(X).sShortLocation, vbTextCompare) = 0 Then   'If we find location..
                iFound = X
                If aConnection.Direction = Outgoing Then
                    If g_aPrograms(X).iAccess = 1 Then 'If it's not allowed access...
                        CloseConnection aConnection.Row
                        KillProcess aConnection.ProcInfo.lProcID
                        CheckConnection.sName = g_aProgramDescriptions(sTmp)
                        CheckConnection.bBlocked = True
                        Exit Function
                    Else
                        CheckConnection.sName = g_aProgramDescriptions(sTmp)
                        CheckConnection.bBlocked = False
                    End If
                    Exit For
                Else
                    If g_aPrograms(X).iServer = 1 Then  'If it's not allowed access...
                        CloseConnection aConnection.Row
                        KillProcess aConnection.ProcInfo.lProcID
                        CheckConnection.sName = g_aProgramDescriptions(sTmp)
                        CheckConnection.bBlocked = True
                        Exit Function
                    Else
                        CheckConnection.sName = g_aProgramDescriptions(sTmp)
                        CheckConnection.bBlocked = False
                    End If
                    Exit For
                End If
            End If
        Next
        If iFound = -1 Then
            CheckConnection = NewExectuable(aConnection, sName) 'Ask user what to do.
        ElseIf (aConnection.Direction = Outgoing And g_aPrograms(iFound).iAccess = 2) Or (aConnection.Direction = Incoming And g_aPrograms(iFound).iServer = 2) Then
            'If its allowed to move.
            FrmMain.lstPrograms.Icons = FrmMain.ilTray
            If aConnection.Direction = Incoming Then
                FrmMain.lstPrograms.ListItems(FindPrograms(LCase(aConnection.ProcInfo.sPath))).SmallIcon = 11
            Else
                FrmMain.lstPrograms.ListItems(FindPrograms(LCase(aConnection.ProcInfo.sPath))).SmallIcon = 12
            End If
            CheckConnection.bBlocked = False
            CheckConnection.sName = sName
        ElseIf aConnection.bTCP = False And g_aPrograms(iFound).iAccess = 2 Then
            FrmMain.lstPrograms.Icons = FrmMain.ilTray
            If aConnection.Direction = Incoming Then
                FrmMain.lstPrograms.ListItems(FindPrograms(LCase(aConnection.ProcInfo.sPath))).SmallIcon = 11
            Else
                FrmMain.lstPrograms.ListItems(FindPrograms(LCase(aConnection.ProcInfo.sPath))).SmallIcon = 12
            End If
            CheckConnection.bBlocked = False
            CheckConnection.sName = sName
        Else
            CheckConnection = NewExectuable(aConnection, sName, iFound)
        End If
    End If
End Function
Private Function NewExectuable(aConnection As tConnectionType, sName As String, Optional iExists As Integer = -1) As tChecking
    Dim lRet                        As VbMsgBoxResult
    Dim sTmp                        As String
    Dim sPath                       As String
    Dim sProgram                    As tProgram
    Dim iTmp                        As Integer
    Dim frmA                        As New frmAlert
    Dim tmpString                   As String
    Dim Item                        As ListItem
    If aConnection.ProcInfo.lProcID > 0 Then PauseProcess (aConnection.ProcInfo.lProcID)
    Load frmA 'Load Alert Form
    frmA.lblProgram(1).Caption = aConnection.ProcInfo.sPath  'Set Caption of Prog Description
    If aConnection.Direction = Incoming Then 'If incoming...
        If FrmMain.chkName.Value = vbUnchecked Then
            sTmp = NameByAddr(aConnection.sRemoteAddr)
            frmA.lblDESC.Caption = sName & " (" & aConnection.ProcInfo.sPath & ") is allowing a connection from " & IIf(Len(sTmp) > 0, sTmp, aConnection.sRemoteAddr) & " on port " & CStr(aConnection.lRemotePort) & " . Do you want to allow this programto act as a server ?"
        Else
            frmA.lblDESC.Caption = sName & " (" & aConnection.ProcInfo.sPath & ") is allowing a connection from " & aConnection.sRemoteAddr & " on port " & CStr(aConnection.lRemotePort) & " . Do you want to allow this programto act as a server ?"
        End If
        frmA.lblDest(0).Caption = "Origin:"
        frmA.lblDest(1).Caption = IIf(Len(sTmp) > 0, sTmp, aConnection.sRemoteAddr)
        frmA.lblPort(0).Caption = "Remote Port:"
        frmA.lblPort(1).Caption = CStr(aConnection.lRemotePort)
        sTmp = PortDetails(CStr(aConnection.lRemotePort), True, IIf(aConnection.bTCP = True, enPortType.TCP, enPortType.UDP))
        frmA.lblPortDesc(0).Caption = "Trojan Info:"
        If Len(sTmp) > 0 Then
            frmA.lblPortDesc(1).Caption = sTmp
        Else
            frmA.lblPortDesc(1).Caption = "[No Description Available]"
        End If
    Else
        If FrmMain.chkName.Value = vbUnchecked Then
            sTmp = NameByAddr(aConnection.sRemoteAddr)
            frmA.lblDESC.Caption = sName & " (" & aConnection.ProcInfo.sPath & ") is trying to connect to " & IIf(Len(sTmp) > 0, sTmp, aConnection.sRemoteAddr) & " using port " & CStr(aConnection.lRemotePort) & " . Do you want to allow this program to access the network ?"
        Else
            frmA.lblDESC.Caption = sName & " (" & aConnection.ProcInfo.sPath & ") is trying to connect to " & aConnection.sRemoteAddr & " using port " & CStr(aConnection.lRemotePort) & " . Do you want to allow this program to access the network ?"
        End If
        frmA.lblDest(0).Caption = "Destination:"
        frmA.lblDest(1).Caption = IIf(Len(sTmp) > 0, sTmp, aConnection.sRemoteAddr)
        frmA.lblPort(0).Caption = "Remote Port:"
        frmA.lblPort(1).Caption = CStr(aConnection.lRemotePort)
        sTmp = PortDetails(CStr(aConnection.lRemotePort), False, IIf(aConnection.bTCP = True, enPortType.TCP, enPortType.UDP))
        frmA.lblPortDesc(0).Caption = "Service Info:"
        If Len(sTmp) > 0 Then
            frmA.lblPortDesc(1).Caption = sTmp
        Else
            frmA.lblPortDesc(1).Caption = "[No Description Available]"
        End If
    End If
    Call FileIconToPicture(aConnection.ProcInfo.sPath, frmA.Pic32, frmA.imgPic) 'Set Icon on the Alert Form
    frmA.Show vbModal, Me 'Show it modaly.
    MakeOntop frmA.hwnd 'Make it ontop
    If frmA.bWhatToDo = True Then 'If we want to allow it.
        If aConnection.Direction = Incoming Then
            If frmA.bRemember = True Then sProgram.iServer = 2 'Set its access.
        Else
            If frmA.bRemember = True Then sProgram.iAccess = 2 'Set its access.
        End If
        UnPauseProcess aConnection.ProcInfo.lProcID  'Unpause the process.
        NewExectuable.bBlocked = False
    Else
        If aConnection.Direction = Incoming Then
            If frmA.bRemember = True Then sProgram.iServer = 1
        Else
            If frmA.bRemember = True Then sProgram.iAccess = 1
        End If
        CloseConnection aConnection.Row
        KillProcess aConnection.ProcInfo.lProcID  'Kill the process.
        NewExectuable.bBlocked = True
    End If
    sProgram.iServer = 0
    sProgram.sLocation = LCase(aConnection.ProcInfo.sPath)
    sProgram.sName = sName
ReFind:
    If iExists = -1 Then
        iTmp = T_UBound(g_aPrograms) + 1
        ReDim Preserve g_aPrograms(iTmp)
    Else
        iTmp = iExists
    End If
    sProgram.iID = iTmp
    g_aPrograms(iTmp) = sProgram
    With sProgram
        sPath = CStr(iTmp + 1)
        sPath = "Software\EliteProdigy\Fire Gate\Programs\" & sPath
        If frmA.bRemember = True Then
            Call REGSaveSetting(vHKEY_LOCAL_MACHINE, sPath, "Name", .sName)
            Call REGSaveSetting(vHKEY_LOCAL_MACHINE, sPath, "Path", .sLocation)
            Call REGSaveSetting(vHKEY_LOCAL_MACHINE, sPath, "Short Path", GetShortPath(.sLocation))
            Call REGSaveSetting(vHKEY_LOCAL_MACHINE, sPath, "ID", CStr(.iID))
            Call REGSaveSetting(vHKEY_LOCAL_MACHINE, sPath, "Access", CStr(.iAccess))
            Call REGSaveSetting(vHKEY_LOCAL_MACHINE, sPath, "Server", CStr(.iServer))
        End If
        If iExists = -1 Then
            g_aProgramDescriptions.Add .sLocation, .sName
            Set Item = FrmMain.lstPrograms.ListItems.Add(, , , , 13)
            Item.ListSubItems.Add , , .sName
            Select Case .iAccess
                Case Is = 0
                    With Item.ListSubItems.Add(, , "Ask")
                        .ForeColor = vbMagenta
                        .Bold = True
                    End With
                Case Is = 1
                    With Item.ListSubItems.Add(, , "Deny")
                        .ForeColor = vbRed
                        .Bold = True
                    End With
                Case Is = 2
                    With Item.ListSubItems.Add(, , "Allow")
                        .ForeColor = vbGreen
                        .Bold = True
                    End With
            End Select
            Select Case .iServer
                Case Is = 0
                    With Item.ListSubItems.Add(, , "Ask")
                        .ForeColor = vbMagenta
                        .Bold = True
                    End With
                Case Is = 1
                    With Item.ListSubItems.Add(, , "Deny")
                        .ForeColor = vbRed
                        .Bold = True
                    End With
                Case Is = 2
                    With Item.ListSubItems.Add(, , "Allow")
                        .ForeColor = vbGreen
                        .Bold = True
                    End With
            End Select
            Item.key = .sLocation
        Else
            If FindPrograms(LCase(.sLocation)) = -1 Then
                iExists = -1
                GoTo ReFind
            End If
            Set Item = FrmMain.lstPrograms.ListItems(FindPrograms(LCase(.sLocation)))
            Select Case .iAccess
                Case Is = 0
                    Item.ListSubItems(2).Text = "Ask"
                    Item.ListSubItems(2).ForeColor = vbMagenta
                Case Is = 1
                    Item.ListSubItems(2).Text = "Deny"
                    Item.ListSubItems(2).ForeColor = vbRed
                Case Is = 2
                    Item.ListSubItems(2).Text = "Allow"
                    Item.ListSubItems(2).ForeColor = vbGreen
            End Select
            Select Case .iServer
                Case Is = 0
                    Item.ListSubItems(3).Text = "Ask"
                    Item.ListSubItems(3).ForeColor = vbMagenta
                Case Is = 1
                    Item.ListSubItems(3).Text = "Deny"
                    Item.ListSubItems(3).ForeColor = vbRed
                Case Is = 2
                    Item.ListSubItems(3).Text = "Allow"
                    Item.ListSubItems(3).ForeColor = vbGreen
            End Select
        End If
    End With
    NewExectuable.sName = sName
    Unload frmA
    Set frmA = Nothing
End Function
