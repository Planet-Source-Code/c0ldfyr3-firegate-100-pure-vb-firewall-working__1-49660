Attribute VB_Name = "ModRegistry"
Option Explicit
Private Declare Function RegOpenKeyEx Lib "advapi32.dll" Alias "RegOpenKeyExA" (ByVal HKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, phkResult As Long) As Long
Private Declare Function RegCloseKey Lib "advapi32.dll" (ByVal HKey As Long) As Long
Private Declare Function RegQueryValueEx Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal HKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, lpData As Any, lpcbData As Long) As Long
Private Declare Function RegSetValueEx Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal HKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, lpData As Any, ByVal cbData As Long) As Long
Private Declare Function RegCreateKeyEx Lib "advapi32.dll" Alias "RegCreateKeyExA" (ByVal HKey As Long, ByVal lpSubKey As String, ByVal Reserved As Long, ByVal lpClass As String, ByVal dwOptions As Long, ByVal samDesired As Long, lpSecurityAttributes As SECURITY_ATTRIBUTES, phkResult As Long, lpdwDisposition As Long) As Long
Private Declare Function RegDeleteKey Lib "advapi32.dll" Alias "RegDeleteKeyA" (ByVal HKey As Long, ByVal lpSubKey As String) As Long
Private Declare Function RegDeleteValue Lib "advapi32.dll" Alias "RegDeleteValueA" (ByVal HKey As Long, ByVal lpValueName As String) As Long
Private Declare Function RegEnumKeyEx Lib "advapi32.dll" Alias "RegEnumKeyExA" (ByVal HKey As Long, ByVal dwIndex As Long, ByVal lpName As String, lpcbName As Long, ByVal lpReserved As Long, ByVal lpClass As String, lpcbClass As Long, lpftLastWriteTime As FILETIME) As Long
Private Declare Function RegEnumValue Lib "advapi32.dll" Alias "RegEnumValueA" (ByVal HKey As Long, ByVal dwIndex As Long, ByVal lpValueName As String, lpcbValueName As Long, ByVal lpReserved As Long, lpType As Long, lpData As Byte, lpcbData As Long) As Long
Private Declare Function RegConnectRegistry Lib "advapi32.dll" Alias "RegConnectRegistryA" (ByVal lpMachineName As String, ByVal HKey As Long, phkResult As Long) As Long
Private Declare Function ExpandEnvironmentStrings Lib "kernel32" Alias "ExpandEnvironmentStringsA" (ByVal lpSrc As String, ByVal lpDst As String, ByVal nSize As Long) As Long
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (dest As Any, Source As Any, ByVal numBytes As Long)
Private Const HKEY_CLASSES_ROOT         As Long = &H80000000
Private Const HKEY_LOCAL_MACHINE        As Long = &H80000002
Private Const HKEY_USERS                As Long = &H80000003
Private Const HKEY_CURRENT_USER         As Long = &H80000001
Private Const REG_OPTION_NON_VOLATILE   As Long = 0
Private Const SYNCHRONIZE               As Long = &H100000
Private Const STANDARD_RIGHTS_ALL       As Long = &H1F0000
Private Const KEY_QUERY_VALUE           As Long = &H1
Private Const KEY_SET_VALUE             As Long = &H2
Private Const KEY_CREATE_SUB_KEY        As Long = &H4
Private Const KEY_ENUMERATE_SUB_KEYS    As Long = &H8
Private Const KEY_NOTIFY                As Long = &H10
Private Const KEY_CREATE_LINK           As Long = &H20
Private Const KEY_ALL_ACCESS            As Long = ((STANDARD_RIGHTS_ALL Or KEY_QUERY_VALUE Or KEY_SET_VALUE Or KEY_CREATE_SUB_KEY Or KEY_ENUMERATE_SUB_KEYS Or KEY_NOTIFY Or KEY_CREATE_LINK) And (Not SYNCHRONIZE))
Private Const ERROR_SUCCESS             As Long = 0&
Private Const ERROR_MORE_DATA           As Long = 234
Private Const ERROR_NO_MORE_ITEMS       As Long = &H103
Private Const ERROR_KEY_NOT_FOUND       As Long = &H2
Enum DataType
    REG_SZ = &H1
    REG_EXPAND_SZ = &H2
    REG_BINARY = &H3
    REG_DWORD = &H4
    REG_MULTI_SZ = &H7
End Enum
Private Const READ_CONTROL = &H20000
Private Const STANDARD_RIGHTS_READ = (READ_CONTROL)
Private Const KEY_READ = ((STANDARD_RIGHTS_READ Or KEY_QUERY_VALUE Or KEY_ENUMERATE_SUB_KEYS Or KEY_NOTIFY) And (Not SYNCHRONIZE))
Private Const KEY_EXECUTE = ((KEY_READ) And (Not SYNCHRONIZE))
Private Type SECURITY_ATTRIBUTES
    nLength                         As Long
    lpSecurityDescriptor            As Long
    bInheritHandle                  As Long
End Type
Private Type FILETIME
    dwLowDateTime                   As Long
    dwHighDateTime                  As Long
End Type
Enum HKEYS
    vHKEY_CLASSES_ROOT = &H80000000
    vHKEY_CURRENT_USER = &H80000001
    vHKEY_LOCAL_MACHINE = &H80000002
    vHKEY_USERS = &H80000003
    vHKEY_PERFORMcANCE_DATA = &H80000004
    vHKEY_CURRENT_CONFIG = &H80000005
    vHKEY_DYN_DATA = &H80000006
End Enum
Dim Security                    As SECURITY_ATTRIBUTES
Public Function REGDeleteSetting(ByVal regHKEY As HKEYS, ByVal sSection As String, Optional ByVal sKey As String) As Boolean
    Dim lReturn                     As Long
    Dim HKey                        As Long
    If Len(sKey) Then
        lReturn = RegOpenKeyEx(regHKEY, REGSubKey(sSection), 0&, KEY_ALL_ACCESS, HKey)
        If lReturn = ERROR_SUCCESS Then
            If sKey = "*" Then sKey = vbNullString
            lReturn = RegDeleteValue(HKey, sKey)
        End If
    Else
        lReturn = RegOpenKeyEx(regHKEY, REGSubKey(), 0&, KEY_ALL_ACCESS, HKey)
        If lReturn = ERROR_SUCCESS Then
            lReturn = RegDeleteKey(HKey, sSection)
        End If
    End If
    REGDeleteSetting = (lReturn = ERROR_SUCCESS)
End Function
Public Function iRegGetSetting(ByVal regHKEY As HKEYS, ByVal sSection As String, ByVal sKey As String, Optional ByVal iDefault As Integer) As Integer
    Dim lReturn                     As Long
    Dim HKey                        As Long
    Dim lType                       As Long
    Dim lBytes                      As Long
    Dim sBuffer                     As String
    iRegGetSetting = iDefault
    lReturn = RegOpenKeyEx(regHKEY, REGSubKey(sSection), 0&, KEY_ALL_ACCESS, HKey)
    If lReturn = 5 Then lReturn = RegOpenKeyEx(regHKEY, REGSubKey(sSection), 0&, KEY_EXECUTE, HKey)
    If lReturn = ERROR_SUCCESS Then
        If sKey = "*" Then
            sKey = vbNullString
        End If
        lReturn = RegQueryValueEx(HKey, sKey, 0&, lType, ByVal sBuffer, lBytes)
        If lReturn = ERROR_SUCCESS Then
            If lBytes > 0 Then
                sBuffer = Space$(lBytes)
                lReturn = RegQueryValueEx(HKey, sKey, 0&, lType, ByVal sBuffer, Len(sBuffer))
                If lReturn = ERROR_SUCCESS Then
                    sBuffer = Left$(sBuffer, lBytes - 1)
                    If IsNumeric(sBuffer) Then iRegGetSetting = CInt(sBuffer)
                End If
            End If
        End If
    End If
End Function
Public Function REGGetSetting(ByVal regHKEY As HKEYS, ByVal sSection As String, ByVal sKey As String, Optional ByVal sDefault As String) As String
    Dim lReturn                     As Long
    Dim HKey                        As Long
    Dim lType                       As Long
    Dim lBytes                      As Long
    Dim sBuffer                     As String
    REGGetSetting = sDefault
    lReturn = RegOpenKeyEx(regHKEY, REGSubKey(sSection), 0&, KEY_ALL_ACCESS, HKey)
    If lReturn = 5 Then
        lReturn = RegOpenKeyEx(regHKEY, REGSubKey(sSection), 0&, KEY_EXECUTE, HKey)
    End If
    If lReturn = ERROR_SUCCESS Then
        If sKey = "*" Then
            sKey = vbNullString
        End If
        lReturn = RegQueryValueEx(HKey, sKey, 0&, lType, ByVal sBuffer, lBytes)
        If lReturn = ERROR_SUCCESS Then
            If lBytes > 0 Then
                sBuffer = Space$(lBytes)
                lReturn = RegQueryValueEx(HKey, sKey, 0&, lType, ByVal sBuffer, Len(sBuffer))
                If lReturn = ERROR_SUCCESS Then
                    REGGetSetting = Left$(sBuffer, lBytes - 1)
                End If
            End If
        End If
    End If
End Function
Public Function REGSaveSetting(ByVal regHKEY As HKEYS, ByVal sSection As String, ByVal sKey As String, ByVal sValue As String) As Boolean
    Dim lRet                        As Long
    Dim HKey                        As Long
    Dim lResult                     As Long
    lRet = RegCreateKeyEx(regHKEY, REGSubKey(sSection), 0&, vbNullString, REG_OPTION_NON_VOLATILE, KEY_ALL_ACCESS, Security, HKey, lResult)
    If lRet = ERROR_SUCCESS Then
        If sKey = "*" Then sKey = vbNullString
        lRet = RegSetValueEx(HKey, sKey, 0&, REG_SZ, ByVal sValue, Len(sValue))
        Call RegCloseKey(HKey)
    End If
    REGSaveSetting = (lRet = ERROR_SUCCESS)
End Function
Private Function REGSubKey(Optional ByVal sSection As String) As String
    If Left$(sSection, 1) = "\" Then
        sSection = Mid$(sSection, 2)
    End If
    If Right$(sSection, 1) = "\" Then
        sSection = Mid$(sSection, 1, Len(sSection) - 1)
    End If
    REGSubKey = sSection
End Function
Private Sub CreateNewKey(sNewKeyName As String, lPredefinedKey As Long)
    Dim HKey                        As Long
    Dim r                           As Long
    r = RegCreateKeyEx(lPredefinedKey, sNewKeyName, 0&, vbNullString, REG_OPTION_NON_VOLATILE, KEY_ALL_ACCESS, Security, HKey, r)
    Call RegCloseKey(HKey)
End Sub
Private Function SetValueEx(ByVal HKey As Long, sValueName As String, lType As Long, vValue As Variant) As Long
    Dim nValue                      As Long
    Dim sValue                      As String
    Select Case lType
        Case REG_SZ
            sValue = vValue & Chr$(0)
            SetValueEx = RegSetValueEx(HKey, sValueName, 0&, lType, sValue, Len(sValue))
        Case REG_DWORD
            nValue = vValue
            SetValueEx = RegSetValueEx(HKey, sValueName, 0&, lType, nValue, 4)
    End Select
End Function
Private Sub SetKeyValue(sKeyName As String, sValueName As String, vValueSetting As Variant, lValueType As Long)
    Dim r                           As Long
    Dim HKey                        As Long
    r = RegOpenKeyEx(HKEY_CLASSES_ROOT, sKeyName, 0, KEY_ALL_ACCESS, HKey)
    r = SetValueEx(HKey, sValueName, lValueType, vValueSetting)
    Call RegCloseKey(HKey)
End Sub
Public Function EnumRegistryKeys(ByVal HKey As HKEYS, ByVal KeyName As String) As Collection
    Dim Handle                      As Long
    Dim Length                      As Long
    Dim Index                       As Long
    Dim subkeyName                  As String
    Dim fFiletime                   As FILETIME
    Set EnumRegistryKeys = New Collection
    If Len(KeyName) Then
        If RegOpenKeyEx(HKey, KeyName, 0, KEY_READ, Handle) Then Exit Function
        HKey = Handle
    End If
    Do
        If cGetInputState(QS_ALLEVENTS) <> 0 Then DoEvents
        Length = 260
        subkeyName = Space$(Length)
        If RegEnumKeyEx(HKey, Index, subkeyName, Length, 0, "", vbNull, fFiletime) = ERROR_NO_MORE_ITEMS Then Exit Do
        subkeyName = Left$(subkeyName, InStr(subkeyName, vbNullChar) - 1)
        EnumRegistryKeys.Add subkeyName, subkeyName
        Index = Index + 1
    Loop
    If Handle Then RegCloseKey Handle
End Function
Public Function EnumRegistryValues(ByVal HKey As HKEYS, ByVal KeyName As String) As Collection
    Dim Handle                      As Long
    Dim Index                       As Long
    Dim valueType                   As Long
    Dim Name                        As String
    Dim namelen                     As Long
    Dim resLong                     As Long
    Dim resString                   As String
    Dim Length                      As Long
    Dim valueInfo(0 To 1)           As Variant
    Dim RetVal                      As Long
    Dim i                           As Integer
    Dim vTemp                       As Variant
    Set EnumRegistryValues = New Collection
    If Len(KeyName) Then
        If RegOpenKeyEx(HKey, KeyName, 0, KEY_READ, Handle) Then Exit Function
        HKey = Handle
    End If
    Do
        namelen = 260
        Name = Space$(namelen)
        Length = 4096
        ReDim resBinary(0 To Length - 1) As Byte
        RetVal = RegEnumValue(HKey, Index, Name, namelen, ByVal 0&, valueType, resBinary(0), Length)
        If RetVal = ERROR_MORE_DATA Then
            ReDim resBinary(0 To Length - 1) As Byte
            RetVal = RegEnumValue(HKey, Index, Name, namelen, ByVal 0&, valueType, resBinary(0), Length)
        End If
        If RetVal Then Exit Do
        valueInfo(0) = Left$(Name, namelen)
        Select Case valueType
            Case REG_DWORD
                CopyMemory resLong, resBinary(0), 4
                valueInfo(1) = resLong
            Case REG_SZ
                If Length <> 0 Then
                    resString = Space$(Length - 1)
                    CopyMemory ByVal resString, resBinary(0), Length - 1
                    valueInfo(1) = resString
                Else
                    valueInfo(1) = ""
                End If
            Case REG_EXPAND_SZ
                If Length <> 0 Then
                    resString = Space$(Length - 1)
                    CopyMemory ByVal resString, resBinary(0), Length - 1
                    Length = ExpandEnvironmentStrings(resString, resString, Len(resString))
                    valueInfo(1) = TrimNull(resString)
                Else
                    valueInfo(1) = ""
                End If
            Case REG_BINARY
                If Length < UBound(resBinary) + 1 Then
                    ReDim Preserve resBinary(0 To Length - 1) As Byte
                End If
                    For i = 0 To UBound(resBinary)
                         resString = resString & " " & Format(Trim(Hex(resBinary(i))), "0#")
                    Next i
                    valueInfo(1) = LTrim(resString)
            Case REG_MULTI_SZ
                resString = Space$(Length - 2)
                CopyMemory ByVal resString, resBinary(0), Length - 2
                resString = Replace(resString, vbNullChar, ",", , , vbBinaryCompare)
                valueInfo(1) = resString
        End Select
        EnumRegistryValues.Add valueInfo, valueInfo(0)
        Index = Index + 1
        If cGetInputState(QS_ALLEVENTS) <> 0 Then DoEvents
    Loop
    If Handle Then RegCloseKey Handle
End Function
Public Function TrimNull(Item As String) As String
    Dim Pos                         As Integer
    Pos = InStr(Item, Chr$(0))
    If Pos Then Item = Left$(Item, Pos - 1)
    TrimNull = Item
End Function


