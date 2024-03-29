VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsHelpCallBack"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''
''  HelpCallBack Class v1.3
''
''  Copyright Olivier Martin 2000
''
''  martin.olivier@bigfoot.com
''
''  This class is used to generate the help events.
''  By creating one in a form module, it's easy
''  to react to its events after passing the object
''  to the install function so CoolMenu can use it.
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''
Option Explicit
Public Event MenuHelp(ByVal MenuText As String, ByVal MenuHelp As String, ByVal Enabled As Boolean)
Public Sub RaiseHelpEvent(ByVal MenuText As String, ByVal MenuHelp As String, Enabled As Boolean)
    RaiseEvent MenuHelp(MenuText, MenuHelp, Enabled)
End Sub
