VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsDictAuditMgr"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Enum AuditTypeEnum
   AuditType_CheckOut = 10
   AuditType_CheckInNew = 13
   AuditType_CheckIn = 15
   AuditType_Unlocked = 20
   
   AuditType_SoundChanged = 54
   AuditType_InfoChanged = 55
   
   AuditType_Transcribed = 60
   
   AuditType_SoundDeleted = 90
End Enum

Public Sub Insert(DictId As Long, AuditType As Integer, DictStatus As Integer)

   If Client.SysSettings.DictationAuditEnable Then
      Dim A As clsDictAudit
      
      Set A = New clsDictAudit
      A.DictId = DictId
      A.AuditType = AuditType
      A.DictStatus = DictStatus
      
      A.UserId = Client.User.UserId
      A.StationId = Client.Station.Id
      Client.Server.DictAuditInsert A
      Set A = Nothing
   End If
End Sub
Public Sub CreateList(DictId As Long)

   Client.Server.CreateDictAuditList DictId
End Sub
Public Function GetNext(Audit As clsDictAudit) As Boolean

   GetNext = Client.Server.DictAuditGetNext(Audit)
End Function
