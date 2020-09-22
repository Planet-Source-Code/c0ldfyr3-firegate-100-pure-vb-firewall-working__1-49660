Attribute VB_Name = "ModWinVer"
Option Explicit
Private Declare Function GetVersionExA Lib "kernel32" (lpVersionInformation As OSVERSIONINFO) As Integer
Private Const PLATFORM_WIN32s        As Integer = 0
Private Const PLATFORM_WIN32_WINDOWS As Integer = 1
Private Const PLATFORM_WIN32_NT      As Integer = 2
Private Type OSVERSIONINFO          'OS Version info, for querying the windows version.
    dwOSVersionInfoSize             As Long
    dwMajorVersion                  As Long
    dwMinorVersion                  As Long
    dwBuildNumber                   As Long
    dwPlatformId                    As Long
    szCSDVersion                    As String * 128
End Type
Public Function GetVersionString() As String
    Dim osinfo                      As OSVERSIONINFO
    Dim retvalue                    As Integer
    osinfo.dwOSVersionInfoSize = 148
    osinfo.szCSDVersion = Space$(128)
    retvalue = GetVersionExA(osinfo) 'Get Version
    With osinfo
        Select Case .dwPlatformId
            Case 1 'If Platform is 9x/Me
                Select Case .dwMinorVersion 'Depends on Minor Version now.
                    Case 0
                        GetVersionString = "Windows 95"
                    Case 10
                        GetVersionString = "Windows 98"
                    Case 90
                        GetVersionString = "Windows Mellinnium"
                End Select
            Case 2 'NT Based
                Select Case .dwMajorVersion 'Depends on Major version this time.
                    Case 3
                        GetVersionString = "Windows NT 3.51"
                    Case 4
                        GetVersionString = "Windows NT 4.0"
                    Case 5
                        If .dwMinorVersion = 0 Then
                            GetVersionString = "Windows 2000"
                        Else
                            GetVersionString = "Windows XP"
                        End If
                End Select
            Case Else
                GetVersionString = "Failed" 'Don't think this should ever happen unless you are using a new windows I haven't heard of yet =)
        End Select
    End With
End Function
Public Function GetVersion() As Long
    Dim osinfo                      As OSVERSIONINFO
    Dim retvalue                    As Integer
    osinfo.dwOSVersionInfoSize = 148
    osinfo.szCSDVersion = Space$(128)
    retvalue = GetVersionExA(osinfo)
    GetVersion = osinfo.dwPlatformId 'Just return the platform ID, good for checking for NT / Win 32 (1 = Win 32 , 0 = NT)
End Function
Public Function IsWinNT() As Boolean
    'Returns True if the current operating system is WinNT
    Dim osvi                        As OSVERSIONINFO
    osvi.dwOSVersionInfoSize = Len(osvi)
    GetVersionExA osvi
    IsWinNT = (osvi.dwPlatformId = PLATFORM_WIN32_NT)
End Function
