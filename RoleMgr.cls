VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsRoleMgr"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Function GetRoles(Roles As clsRoles, ByVal GroupId As Long, ByVal OrgId As Long) As Boolean

   GetRoles = Client.Server.GetRoles(Roles, GroupId, OrgId)
End Function
Public Sub SaveRoles(Roles As clsRoles)

   Dim Eno As Long

   On Error GoTo SaveRoles_Err
   
   Client.Server.SaveRoles Roles
   Client.LoggMgr.Insert 1320110, LoggLevel_SysAdmin, 0, Roles.LoggData
   Exit Sub
   
SaveRoles_Err:
   Eno = Err.Number
   ErrorHandle "1290101", Eno, 1290101, "Rollerna kan inte sparas", False
   Resume Next
End Sub

