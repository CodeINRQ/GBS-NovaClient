VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsOrg"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public OrgId         As Long
Public OrgParent     As Long
Public OrgText       As String
Public DictContainer As Boolean
Public ShowInTree    As Boolean
Public ShowParent    As Long
Public Roles         As clsRoles
Public TreeLevel     As Integer  '0 is root level
Public ShowBelow     As Boolean

Public Property Get LoggData() As String

   LoggData = OrgText & "/" & CStr(OrgId)
End Property
