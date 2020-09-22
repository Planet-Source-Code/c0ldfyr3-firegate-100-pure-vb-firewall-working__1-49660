Attribute VB_Name = "ModPrograms"
Option Explicit
Public Type tProgram
    sName                           As String
    sLocation                       As String
    sShortLocation                  As String
    iServer                         As Integer
    iAccess                         As Integer
    iID                             As Integer
End Type
Public Sub LoadPrograms()
    Static bLoaded                  As Boolean
    Dim X                           As Integer
    Dim sCol                        As Collection
    Dim sProgram                    As tProgram
    Dim sTmp                        As String
    Dim Item                        As ListItem
    Debug.Print "Loading"
    If bLoaded = True Then Exit Sub
    bLoaded = True
    Set sCol = EnumRegistryKeys(vHKEY_LOCAL_MACHINE, "Software\EliteProdigy\Fire Gate\Programs")
    LoadProcessUserIDs
    GetProcs
    If sCol.Count > 0 Then
        ReDim g_aPrograms(sCol.Count - 1)
        For X = 1 To sCol.Count
            With sProgram
                sTmp = CStr(sCol(X))
                sTmp = "Software\EliteProdigy\Fire Gate\Programs\" & sTmp
                .sName = REGGetSetting(vHKEY_LOCAL_MACHINE, sTmp, "Name")
                .sLocation = LCase(REGGetSetting(vHKEY_LOCAL_MACHINE, sTmp, "Path"))
                .sShortLocation = LCase(REGGetSetting(vHKEY_LOCAL_MACHINE, sTmp, "Short Path"))
                .iID = X - 1
                .iAccess = iRegGetSetting(vHKEY_LOCAL_MACHINE, sTmp, "Access", 0)
                .iServer = iRegGetSetting(vHKEY_LOCAL_MACHINE, sTmp, "Server", 0)
                g_aProgramDescriptions.Add .sLocation, .sName
                FrmMain.lstPrograms.Icons = FrmMain.ilTray
                FrmMain.lstPrograms.SmallIcons = FrmMain.ilTray
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
            End With
            g_aPrograms(X - 1) = sProgram
        Next
    End If
    Set Item = Nothing
End Sub
Public Function T_UBound(aArr() As tProgram) As Integer
    On Error GoTo ErrClear
    T_UBound = UBound(aArr)
    Exit Function
ErrClear:
    Err.Clear
    T_UBound = -1
End Function
