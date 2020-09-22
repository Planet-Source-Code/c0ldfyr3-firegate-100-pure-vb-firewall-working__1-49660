VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsMyItemDatas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''
''  CMyItemDatas Class v1.3
''
''  Copyright Olivier Martin 2000
''
''  martin.olivier@bigfoot.com
''
''  This is a collection of CMyItemData that will
''  old information about converted menus in
''  OwnerDraw.
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''
Option Explicit
Private mCol                        As Collection
Public Function Add(sKey As String) As clsMyItemData
    'create a new object
    Dim objNewMember                As clsMyItemData
    Set objNewMember = New clsMyItemData
    mCol.Add objNewMember, sKey
    'return the object created
    Set Add = objNewMember
    Set objNewMember = Nothing
End Function
Public Property Get Item(vntIndexKey As Variant, Optional Source As String) As clsMyItemData
  On Error GoTo hError
  Set Item = mCol(vntIndexKey)
  Exit Property
hError:
  Debug.Print Source$; " "; Err.Description; " ("; Err.Number; ")"
  Err.Clear
End Property
Public Property Get Count() As Long
    Count = mCol.Count
End Property
Public Sub Remove(vntIndexKey As Variant)
    mCol.Remove vntIndexKey
End Sub
Public Property Get NewEnum() As IUnknown
    Set NewEnum = mCol.[_NewEnum]
End Property
Private Sub Class_Initialize()
    Set mCol = New Collection
End Sub
Private Sub Class_Terminate()
    Set mCol = Nothing
End Sub
