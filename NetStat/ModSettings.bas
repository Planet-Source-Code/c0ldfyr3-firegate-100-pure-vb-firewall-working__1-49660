Attribute VB_Name = "ModSettings"
Option Explicit
Public Sub SaveSettings()
    Dim sPath                       As String
    Dim sTmp                        As String
    sPath = "Software\EliteProdigy\Fire Gate\Settings"
    Call REGSaveSetting(vHKEY_LOCAL_MACHINE, sPath, "Resolve HostNames", IIf(FrmMain.chkName.Value = vbChecked, "1", "0"))
    Call REGSaveSetting(vHKEY_LOCAL_MACHINE, sPath, "Show Process Icons", IIf(FrmMain.chkIcons.Value = vbChecked, "1", "0"))
End Sub
Public Sub LoadSettings()
    Dim sPath                       As String
    Dim iTmp                        As Integer
    sPath = "Software\EliteProdigy\Fire Gate\Settings"
    iTmp = iRegGetSetting(vHKEY_LOCAL_MACHINE, sPath, "Resolve HostNames", 0)
    If iTmp = 1 Then FrmMain.chkName.Value = vbChecked
    iTmp = iRegGetSetting(vHKEY_LOCAL_MACHINE, sPath, "Show Process Icons", 1)
    If iTmp = 1 Then FrmMain.chkIcons.Value = vbChecked
End Sub
Public Sub SavePorts()
    Dim X                           As Integer
    Dim sTmp                        As String
    X = FrmMain.lstPorts.ListItems.Count
    sTmp = "Software\EliteProdigy\Fire Gate\Settings\Ports\"
    Call REGDeleteSetting(vHKEY_LOCAL_MACHINE, sTmp, vbNullString)
    For X = 1 To X
        Call REGSaveSetting(vHKEY_LOCAL_MACHINE, sTmp, FrmMain.lstPorts.ListItems(X).Text, FrmMain.lstPorts.ListItems(X).ListSubItems(1).Tag)
    Next
End Sub
Public Sub SaveIPs()
    Dim X                           As Integer
    Dim sTmp                        As String
    X = FrmMain.lstIPs.ListItems.Count
    sTmp = "Software\EliteProdigy\Fire Gate\Settings\IPS\"
    Call REGDeleteSetting(vHKEY_LOCAL_MACHINE, sTmp, vbNullString)
    For X = 1 To X
        Call REGSaveSetting(vHKEY_LOCAL_MACHINE, sTmp, FrmMain.lstIPs.ListItems(X).Text, FrmMain.lstIPs.ListItems(X).ListSubItems(1).Tag)
    Next
End Sub
Public Sub LoadPorts()
    Dim sCol                        As Collection
    Dim sTmp                        As String
    Dim X                           As Integer
    Set sCol = EnumRegistryValues(vHKEY_LOCAL_MACHINE, "Software\EliteProdigy\Fire Gate\Settings\Ports\")
    For X = 1 To sCol.Count
        If IsNumeric(sCol(X)(0)) Then
            sTmp = CStr(sCol(X)(1))
            Select Case sTmp
                Case Is = "0"
                    sTmp = "Both"
                Case Is = "1"
                    sTmp = "In"
                Case Is = "2"
                    sTmp = "Out"
                Case Else
                    sTmp = ""
            End Select
            If Len(sTmp) > 0 Then FrmMain.lstPorts.ListItems.Add(, , CStr(sCol(X)(0))).ListSubItems.Add(, , sTmp).Tag = CStr(sCol(X)(1))
        End If
    Next
End Sub
Public Sub LoadIPs()
    Dim sCol                        As Collection
    Dim sTmp                        As String
    Dim X                           As Integer
    Set sCol = EnumRegistryValues(vHKEY_LOCAL_MACHINE, "Software\EliteProdigy\Fire Gate\Settings\IPs\")
    For X = 1 To sCol.Count
        sTmp = CStr(sCol(X)(1))
        Select Case sTmp
            Case Is = "0"
                sTmp = "Both"
            Case Is = "1"
                sTmp = "In"
            Case Is = "2"
                sTmp = "Out"
            Case Else
                sTmp = ""
        End Select
        If Len(sTmp) > 0 Then FrmMain.lstIPs.ListItems.Add(, , CStr(sCol(X)(0))).ListSubItems.Add(, , sTmp).Tag = CStr(sCol(X)(1))
    Next
End Sub
